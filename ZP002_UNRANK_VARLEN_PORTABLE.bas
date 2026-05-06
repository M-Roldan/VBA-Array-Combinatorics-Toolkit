Attribute VB_Name = "ZP002_UNRANK_VARLEN_PORTABLE"
Option Explicit

' Double loses integer precision above 2^53.
' Beyond this point, index validation is unreliable.
Private Const P002_MAX_SAFE_DOUBLE As Double = 9.00719925474099E+15 ' 2^53

' ============================================================
' MODULE: ZP002_UNRANK_VARLEN_PORTABLE
' CATEGORY: Combinatorics / Variable-Length Radix
'
' PURPOSE:
' This module provides a self-contained, portable implementation
' of the variable-length radix unranking function
' P002_UNRANK_VARLEN. It is designed to be copied and pasted
' into any VBA project without requiring any other module
' from the library.
'
' ------------------------------------------------------------
' PUBLIC API
' ------------------------------------------------------------
'
'   P002_UNRANK_VARLEN
'     Returns a single combination at a specified 1-based
'     index from a variable-length radix space defined by a
'     set of input values. Higher indices map to longer
'     combinations.
'
' ------------------------------------------------------------
' INTERNAL FUNCTIONS
' ------------------------------------------------------------
'
'   P002_VarlenRadix_Unrank
'     Core engine. Locates the block containing the target
'     index, computes the local index within that block, and
'     decomposes it into per-position digits using direct
'     division. Called once per P002_UNRANK_VARLEN invocation.
'
'   P002_ARR_TO_1D_STRING
'     Normalizes any supported input into a 1-based 1D String
'     array. Rejects error values and non-vector inputs.
'
'   P002_ARR_NDIM
'     Returns the number of dimensions of an array.
'
'   P002_ARR_TO_2D
'     Normalizes any input into a canonical 2D array.
'
'   P002_ARR_SHAPE
'     Returns the size of each dimension as a 1D array.
'
' ------------------------------------------------------------
' DESIGN CONVENTIONS
' ------------------------------------------------------------
'
' 1. SELF-CONTAINED
'    This module has no dependencies on any other module in
'    the library. All required utilities are implemented
'    locally with the P002_ prefix.
'
' 2. ERROR SIGNALING
'    This module does not participate in the centralized error
'    system (LibError enum in ERRORS.bas). All errors are
'    raised via vbObjectError + offset and caught by
'    P002_UNRANK_VARLEN, which returns CVErr(xlErrNum) or
'    CVErr(xlErrValue) as appropriate.
'
'    Local error code offsets:
'      vbObjectError + 1   index out of range or space overflow
'      vbObjectError + 2   input is not a scalar or 1D vector
'      vbObjectError + 3   input contains error values
'
' 3. INDEXING
'    Combinations are numbered 1-based (1 = first combination).
'    Internally, the local block index is converted to
'    zero-based before digit decomposition.
'
' 4. ENUMERATION ORDER
'    Combinations are grouped by length. All 1-element
'    combinations come first, then all 2-element combinations,
'    and so on. Within each group, combinations follow the
'    radix order defined by the position of each value in the
'    input.
'
' ============================================================


' ============================================================
' FUNCTION: P002_UNRANK_VARLEN
' CATEGORY: Combinatorics / Variable-Length Radix
'
' DESCRIPTION:
' The P002_UNRANK_VARLEN function returns the single
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
'    Delegates to P002_ARR_TO_1D_STRING, which converts any
'    supported input into a 1-based 1D String array. Raises
'    vbObjectError + 2 if input is not a scalar or 1D vector.
'    Raises vbObjectError + 3 if input contains error values
'    (e.g. #N/A, #DIV/0!).
'
' 2. Retrieve combination at targetIndex
'    Delegates to P002_VarlenRadix_Unrank, which locates the
'    block containing targetIndex, decomposes the local index
'    using direct division, and maps the result to actual
'    values in valueSet.
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
' A 1-based (1 x comboLen) 2D Variant array representing the
' combination at targetIndex, where comboLen is the length
' of the combination.
' Returns a CVErr on failure:
'   xlErrNum   if input is not a valid 1D vector, if it
'              contains error values, if targetIndex < 1,
'              or if the combination space exceeds Double
'              precision
'   xlErrValue for any other unexpected error
'
' EXAMPLE:
' ' Worksheet usage:
' =P002_UNRANK_VARLEN(4, A1:A2)
' ' Returns the 4th combination of the values in A1:A2
'
' ' VBA usage:
' Dim result As Variant
' result = P002_UNRANK_VARLEN(4, Array("A", "B"))
' ' valueSet = ("A", "B"), combinations in order:
' '   Index 1 --> ("A")
' '   Index 2 --> ("B")
' '   Index 3 --> ("A", "A")
' '   Index 4 --> ("A", "B")  <-- returns this
'
' ============================================================
Public Function P002_UNRANK_VARLEN( _
    ByVal targetIndex As Long, _
    ByVal inputs As Variant _
) As Variant

    On Error GoTo HandleError

    Dim valueSet As Variant

    ' --------------------------------------------------------
    ' STEP 1: Normalize input --> valueSet
    ' Raises vbObjectError + 2 if input is not a scalar or
    ' 1D vector. Raises vbObjectError + 3 if input contains
    ' error values (e.g. #N/A, #DIV/0!).
    ' --------------------------------------------------------
    valueSet = P002_ARR_TO_1D_STRING(inputs)

    ' --------------------------------------------------------
    ' STEP 2: Retrieve combination at targetIndex
    ' Raises vbObjectError + 1 if targetIndex < 1 or if the
    ' combination space exceeds Double precision.
    ' --------------------------------------------------------
    P002_UNRANK_VARLEN = P002_VarlenRadix_Unrank(valueSet, targetIndex)

    Exit Function

HandleError:

    Select Case Err.Number
        Case vbObjectError + 1, vbObjectError + 2, vbObjectError + 3
            P002_UNRANK_VARLEN = CVErr(xlErrNum)
        Case Else
            P002_UNRANK_VARLEN = CVErr(xlErrValue)
    End Select

End Function


' ============================================================
' FUNCTION: P002_VarlenRadix_Unrank
' CATEGORY: Combinatorics / Variable-Length Radix (Internal)
'
' DESCRIPTION:
' The P002_VarlenRadix_Unrank function returns the single
' combination at a given 1-based index from the variable-
' length radix space defined by valueSet, without computing
' the full enumeration.
'
' Unlike fixed-length Cartesian unranking, the output length
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
' The digit decomposition is implemented directly via
' successive division, without delegating to the Radix Engine.
' This keeps the module self-contained and avoids porting the
' full Radix Engine.
'
' PRECONDITIONS (caller must guarantee):
'   valueSet must be a non-empty 1-based 1D String array.
'   targetIndex must be >= 1.
'   Use P002_ARR_TO_1D_STRING to normalize inputs before
'   calling.
'
' STRATEGY:
' 1. Compute the size of valueSet
'    The number of available values (valueSetSize) determines
'    the base of the radix system and the block sizes.
'    Exits silently if valueSet is empty.
'
' 2. Validate the lower index bound
'    Raises vbObjectError + 1 if targetIndex < 1.
'
' 3. Locate the block containing targetIndex
'    Combinations are grouped into blocks by output length:
'    block 1 has N combinations (length 1), block 2 has N^2
'    (length 2), etc. The loop advances through blocks until
'    the running total reaches or exceeds targetIndex. An
'    overflow guard raises vbObjectError + 1 if the space
'    exceeds Double precision before targetIndex is reached.
'
' 4. Compute the local index within the block (1-based)
'    Subtracts the cumulative total of all preceding blocks.
'
' 5. Convert local index to zero-based
'    Subtracts 1 before digit decomposition.
'
' 6. Decompose zero-based index into per-position digits
'    Applies successive division by valueSetSize from LSD to
'    MSD. Each remainder is the 0-based position of that
'    output column within valueSet.
'
' 7. Map digits to actual values in valueSet
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
Private Function P002_VarlenRadix_Unrank( _
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
        Err.Raise vbObjectError + 1, , "Index out of range"
    End If

    ' --------------------------------------------------------
    ' STEP 3: Locate the block containing targetIndex
    ' Combinations are grouped by output length. Block k
    ' contains all combinations of length k (valueSetSize^k
    ' combinations). The loop advances until the running total
    ' reaches or exceeds targetIndex.
    ' An overflow guard raises vbObjectError + 1 if the space
    ' exceeds Double precision before targetIndex is reached.
    ' --------------------------------------------------------
    Dim comboLen As Long:          comboLen = 1
    Dim prevBlocksTotal As Double: prevBlocksTotal = 0
    Dim blockSize As Double:       blockSize = 1
    Dim runningTotal As Double

    Do While targetIndex > runningTotal
        If blockSize > P002_MAX_SAFE_DOUBLE / valueSetSize Then
            Err.Raise vbObjectError + 1, , _
                "Combination space exceeds Double precision limit"
        End If
        blockSize = blockSize * valueSetSize
        runningTotal = blockSize + prevBlocksTotal
        prevBlocksTotal = runningTotal
        comboLen = comboLen + 1
    Loop

    ' --------------------------------------------------------
    ' STEP 4: Compute local index within the block (1-based)
    ' Subtracts the cumulative total of all preceding blocks.
    ' --------------------------------------------------------
    Dim blockIndex As Long
    blockIndex = targetIndex - (runningTotal - blockSize)

    ' --------------------------------------------------------
    ' STEP 5: Convert local index to zero-based
    ' Digit decomposition expects a zero-based value since
    ' digit positions start at 0.
    ' --------------------------------------------------------
    Dim zeroIndex As Long
    zeroIndex = blockIndex - 1

    ' --------------------------------------------------------
    ' STEP 6: Decompose zero-based index into digits
    ' Applies successive division by valueSetSize from LSD
    ' to MSD. Each remainder is the 0-based position of that
    ' output column within valueSet.
    ' --------------------------------------------------------
    Dim outputLen As Long
    outputLen = comboLen - 1

    Dim digits() As Long
    ReDim digits(1 To outputLen)

    Dim j As Long
    Dim remaining As Long
    remaining = zeroIndex

    For j = outputLen To 1 Step -1
        digits(j) = remaining Mod valueSetSize
        remaining = remaining \ valueSetSize
    Next j

    ' --------------------------------------------------------
    ' STEP 7: Map digits to actual values in valueSet
    ' Each digit is the 0-based position within valueSet.
    ' The lb offset makes the lookup safe for any lower bound.
    ' --------------------------------------------------------
    Dim result() As Variant
    ReDim result(1 To 1, 1 To outputLen)

    Dim dimLB As Long
    dimLB = LBound(valueSet)

    Dim i As Long
    For i = 1 To outputLen
        result(1, i) = valueSet(dimLB + digits(i))
    Next i

    P002_VarlenRadix_Unrank = result

End Function


' ============================================================
' FUNCTION: P002_ARR_TO_1D_STRING
' CATEGORY: Array Utilities / Normalization (Internal)
'
' DESCRIPTION:
' The P002_ARR_TO_1D_STRING function converts any supported
' input into a 1-based 1D array of String values.
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
'    Calls P002_ARR_TO_2D so that all inputs share a common
'    structural representation.
'
' 2. Retrieve the shape
'    Rows and columns are obtained via P002_ARR_SHAPE.
'
' 3. Dispatch by shape
'    The function branches on the shape of the normalized
'    input:
'      • (1 x N) --> row vector case (covers scalar when N = 1)
'      • (N x 1) --> column vector case
'      • otherwise --> raise vbObjectError + 2
'
' 4. Copy values with CStr coercion
'    Each element is validated against IsError before
'    conversion. Raises vbObjectError + 3 if an error value
'    is found.
'
' RETURN VALUE:
' A 1-based 1D String array.
'
' ============================================================
Private Function P002_ARR_TO_1D_STRING(ByVal inputValue As Variant) As String()

    Dim result() As String
    Dim i As Long

    Dim arr As Variant
    Dim shape As Variant

    ' --------------------------------------------------------
    ' Normalize to 2D array always
    ' --------------------------------------------------------
    arr = P002_ARR_TO_2D(inputValue)

    shape = P002_ARR_SHAPE(arr)

    Dim rows As Long, cols As Long
    rows = shape(1)
    cols = shape(2)

    ' --------------------------------------------------------
    ' CASE 1: Row vector (covers scalar when cols = 1)
    ' --------------------------------------------------------
    If rows = 1 Then
        ReDim result(1 To cols)

        For i = 1 To cols
            If IsError(arr(1, i)) Then Err.Raise vbObjectError + 3, , _
                "Array contains error value"
            result(i) = CStr(arr(1, i))
        Next i

        P002_ARR_TO_1D_STRING = result
        Exit Function
    End If

    ' --------------------------------------------------------
    ' CASE 2: Column vector
    ' --------------------------------------------------------
    If cols = 1 Then
        ReDim result(1 To rows)

        For i = 1 To rows
            If IsError(arr(i, 1)) Then Err.Raise vbObjectError + 3, , _
                "Array contains error value"
            result(i) = CStr(arr(i, 1))
        Next i

        P002_ARR_TO_1D_STRING = result
        Exit Function
    End If

    Err.Raise vbObjectError + 2, , "Value must be scalar or 1D"

End Function


' ============================================================
' FUNCTION: P002_ARR_NDIM
' CATEGORY: Array Utilities / Inspection (Internal)
'
' DESCRIPTION:
' The P002_ARR_NDIM function returns the number of dimensions
' of a VBA array. Returns 0 for non-arrays and scalars.
' If the input is a Range, its .Value is extracted before
' analysis.
'
' ============================================================
Private Function P002_ARR_NDIM(ByVal inputArray As Variant) As Long

    If TypeName(inputArray) = "Range" Then
        inputArray = inputArray.value
    End If

    ' --------------------------------------------------------
    ' CASE 1: Not an array
    ' --------------------------------------------------------
    If Not IsArray(inputArray) Then
        P002_ARR_NDIM = 0
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Probe each dimension safely.
    ' Stop when UBound raises an error (dimension does not
    ' exist). Upper probe limit is 60 (VBA maximum).
    ' --------------------------------------------------------
    Dim dimensionIndex As Long

    On Error GoTo ExitFunction

    For dimensionIndex = 1 To 60
        Dim tmp As Long
        tmp = UBound(inputArray, dimensionIndex)
        P002_ARR_NDIM = dimensionIndex
    Next dimensionIndex

ExitFunction:
    On Error GoTo 0

End Function


' ============================================================
' FUNCTION: P002_ARR_TO_2D
' CATEGORY: Array Utilities / Normalization (Internal)
'
' DESCRIPTION:
' The P002_ARR_TO_2D function normalizes any supported input
' into a 2D array with 1-based indexing.
'
'   • Range  --> .Value (always 2D from Excel)
'   • Scalar --> (1 x 1) 2D array
'   • 1D array --> (N x 1) column
'   • 2D array --> returned as-is
'
' ============================================================
Private Function P002_ARR_TO_2D(ByVal inputValue As Variant) As Variant

    ' --------------------------------------------------------
    ' CASE 1: Input is a Range
    ' --------------------------------------------------------
    If TypeName(inputValue) = "Range" Then
        Dim rangeVal As Variant
        rangeVal = inputValue.value
        If IsArray(rangeVal) Then
            P002_ARR_TO_2D = rangeVal
        Else
            Dim singleCell(1 To 1, 1 To 1) As Variant
            singleCell(1, 1) = rangeVal
            P002_ARR_TO_2D = singleCell
        End If
        Exit Function
    End If

    ' --------------------------------------------------------
    ' CASE 2: Input is already an array
    ' --------------------------------------------------------
    If IsArray(inputValue) Then

        ' ----------------------------------------------------
        ' Subcase: 1D array --> reshape to (N x 1) column
        ' ----------------------------------------------------
        If P002_ARR_NDIM(inputValue) = 1 Then

            Dim i As Long
            Dim result() As Variant

            ReDim result(1 To UBound(inputValue) - LBound(inputValue) + 1, 1 To 1)

            For i = LBound(inputValue) To UBound(inputValue)
                result(i - LBound(inputValue) + 1, 1) = inputValue(i)
            Next i

            P002_ARR_TO_2D = result
            Exit Function

        End If

        ' ----------------------------------------------------
        ' Subcase: already 2D (or higher) --> return as-is
        ' ----------------------------------------------------
        P002_ARR_TO_2D = inputValue
        Exit Function

    End If

    ' --------------------------------------------------------
    ' CASE 3: Scalar --> wrap in (1 x 1) 2D array
    ' --------------------------------------------------------
    Dim temp(1 To 1, 1 To 1) As Variant
    temp(1, 1) = inputValue

    P002_ARR_TO_2D = temp

End Function


' ============================================================
' FUNCTION: P002_ARR_SHAPE
' CATEGORY: Array Utilities / Inspection (Internal)
'
' DESCRIPTION:
' The P002_ARR_SHAPE function returns the size of each
' dimension of a VBA array as a 1-based 1D Long array.
' Returns Empty for non-arrays.
'
' ============================================================
Private Function P002_ARR_SHAPE(arr As Variant) As Variant

    Dim dims As Long
    dims = P002_ARR_NDIM(arr)

    If dims = 0 Then Exit Function

    Dim result() As Long
    ReDim result(1 To dims)

    Dim i As Long
    For i = 1 To dims
        result(i) = UBound(arr, i) - LBound(arr, i) + 1
    Next i

    P002_ARR_SHAPE = result

End Function
