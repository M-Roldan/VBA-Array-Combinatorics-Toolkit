Attribute VB_Name = "MATH"
Option Explicit

#If Win64 Then

' ============================================================
' MODULE: MATH
' CATEGORY: Math / Number Theory / Numeral Systems
'
' PURPOSE:
' This module provides two independent mathematical engines:
' the GCD Engine for greatest common divisor computation, and
' the Radix Engine for base conversion between pure and mixed
' numeral systems.
'
' ------------------------------------------------------------
' PUBLIC API
' ------------------------------------------------------------
'
'   MATH_GCD
'     Computes the GCD of a variable number of integer inputs.
'     Accepts any combination of scalars, ranges, and arrays.
'     Requires Office 64-bit (Win64).
'
'   MATH_RADIX_CONVERT
'     Converts a number from any base (pure or mixed) to any
'     other base. Primary interface of the Radix Engine.
'
'   MATH_RADIX_TO_DECIMAL
'     Converts a digit array from any base to decimal.
'     Exposed as public for direct use by advanced callers.
'
'   MATH_RADIX_FROM_DECIMAL
'     Converts a decimal value to any base representation.
'     Exposed as public for direct use by advanced callers.
'
' ------------------------------------------------------------
' GCD ENGINE
' ------------------------------------------------------------
'
'   MATH_GCD            -- public UDF entry point
'   GCD_Two             -- core Euclidean algorithm (two values)
'   GCD_Try_To_LongLong -- safe numeric coercion to LongLong
'   GCD_Count_Numeric   -- first-pass counter (exact allocation)
'   GCD_Append_Numeric  -- second-pass collector
'
'   Requires Office 64-bit (Win64). The 32-bit fallback
'   returns #N/A and does nothing else.
'
' ------------------------------------------------------------
' RADIX ENGINE
' ------------------------------------------------------------
'
'   MATH_RADIX_CONVERT        -- public entry point
'   MATH_RADIX_TO_DECIMAL     -- public conversion primitive
'   MATH_RADIX_FROM_DECIMAL   -- public conversion primitive
'   Radix_Normalize_Base      -- internal base normalization
'   Radix_Normalize_Digits    -- internal digit normalization
'   Radix_Pure_To_Decimal     -- pure base --> decimal
'   Radix_Mixed_To_Decimal    -- mixed base --> decimal
'   Radix_Decimal_To_Pure     -- decimal --> pure base
'   Radix_Decimal_To_Mixed    -- decimal --> mixed base
'
' ------------------------------------------------------------
' DESIGN CONVENTIONS
' ------------------------------------------------------------
'
' 1. PURE VS MIXED BASE
'    A pure base is a scalar Long >= 2 applied uniformly to
'    all digit positions. A mixed base is a 1D Long array
'    where each element defines the base of one position.
'
' 2. DIGIT ORDER
'    All digit arrays use MSD --> LSD order (most significant
'    digit first), consistent with standard positional
'    notation.
'
' 3. INDEXING
'    All digit arrays produced by this module are 1-based,
'    consistent with the library-wide convention.
'
' 4. ERROR SIGNALING
'    All errors are raised via Err.Raise ERR_*. No public
'    function in this module returns CVErr -- callers are
'    responsible for handling errors at their level.
'
' ------------------------------------------------------------
' DEPENDENCIES
' ------------------------------------------------------------
'
'   ARRAYS module:
'     ARR_TO_2D, ARR_SHAPE, ARR_TO_1D_LONG
'     (used by Radix_Normalize_Base and Radix_Normalize_Digits)
'
'   TEXT module:
'     TXT_NUM_TO_DIGITS
'     (used by Radix_Normalize_Digits to decompose scalar inputs)
'
'   ERRORS module:
'     ERR_RADIX_* codes, ERR_TXT_INVALID_NUMERIC is raised
'     transitively via TXT_NUM_TO_DIGITS
'
' ------------------------------------------------------------
' RELATED MODULES
' ------------------------------------------------------------
'
'   COMBINATORICS module:
'     Uses MATH_RADIX_CONVERT in
'     Build_Cartesian_Product_Math_Index for index
'     decomposition into mixed base digits.
'
' ============================================================







' ============================================================
' FUNCTION: GCD_Two
'
' DESCRIPTION:
' Computes the GCD of two 64-bit integers using the iterative
' Euclidean algorithm. Handles negative inputs by taking their
' absolute value before processing.
'
' PARAMETERS:
'   a   LongLong   First integer. Negative values are accepted.
'   b   LongLong   Second integer. Negative values are accepted.
'
' RETURN VALUE:
' A LongLong representing the GCD of a and b.
'
' EXAMPLE:
' GCD_Two(48, 18)   ' Returns: 6
' GCD_Two(-12, 8)   ' Returns: 4
'
' ============================================================
Private Function GCD_Two(ByVal a As LongLong, ByVal b As LongLong) As LongLong

    ' Ensure non-negative values
    If a < 0 Then a = -a
    If b < 0 Then b = -b

    ' Handle trivial cases
    If a = 0 Then
        GCD_Two = b
        Exit Function
    End If
    If b = 0 Then
        GCD_Two = a
        Exit Function
    End If

    ' Euclidean algorithm (iterative to avoid recursion limits)
    Dim t As LongLong
    Do While b <> 0
        t = a Mod b
        a = b
        b = t
    Loop

    GCD_Two = a

End Function


' ============================================================
' FUNCTION: GCD_Try_To_LongLong
'
' DESCRIPTION:
' Attempts to coerce a Variant into a 64-bit integer (LongLong).
' Returns True and sets dst if successful. Returns False if the
' value is non-numeric, empty, or causes an overflow.
'
' PARAMETERS:
'   src   Variant    The value to coerce. Any Variant is accepted;
'                    non-numeric and overflow cases return False.
'   dst   LongLong   Output parameter. Set only when the function
'                    returns True.
'
' RETURN VALUE:
' Boolean. dst is set only when the return value is True.
'
' ============================================================
Private Function GCD_Try_To_LongLong(ByVal src As Variant, ByRef dst As LongLong) As Boolean

    On Error GoTo Fail

    If IsNumeric(src) Then
        ' Reject empty-like strings
        If Len(Trim$(CStr(src))) = 0 Then GoTo Fail
        dst = CLngLng(src)
        GCD_Try_To_LongLong = True
        Exit Function
    End If

Fail:
    GCD_Try_To_LongLong = False

End Function


' ============================================================
' FUNCTION: GCD_Count_Numeric
'
' DESCRIPTION:
' First-pass companion to GCD_Append_Numeric. Counts how many
' valid LongLong-coercible values exist in src without allocating
' any output array. Used to enable exact allocation before the
' second pass.
'
' PARAMETERS:
'   src   Variant   The value to scan. Accepted forms:
'                     - Range object (single or multi-cell)
'                     - Scalar value
'                     - 1D array
'                     - 2D array
'
' RETURN VALUE:
' A Long representing the count of valid numeric values found.
'
' ============================================================
Private Function GCD_Count_Numeric(ByVal src As Variant) As Long

    Dim count As Long
    Dim tmp As LongLong
    Dim v As Variant

    ' --------------------------------------------------------
    ' Case 1: Range object
    ' --------------------------------------------------------
    If IsObject(src) Then
        Dim rng As Range, cell As Range
        On Error Resume Next
        Set rng = src
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each cell In rng.Cells
                If Not IsError(cell.Value2) Then
                    If GCD_Try_To_LongLong(cell.Value2, tmp) Then
                        count = count + 1
                    End If
                End If
            Next cell
        End If
        GCD_Count_Numeric = count
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Case 2: Scalar
    ' --------------------------------------------------------
    If Not IsArray(src) Then
        If VarType(src) <> vbError Then
            If GCD_Try_To_LongLong(src, tmp) Then count = 1
        End If
        GCD_Count_Numeric = count
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Case 3: 1D array
    ' --------------------------------------------------------
    Dim i As Long
    On Error GoTo Try2D
    Dim lb1 As Long, ub1 As Long
    lb1 = LBound(src, 1)
    ub1 = UBound(src, 1)

    ' Verify it is truly 1D by checking that dimension 2 does not exist
    Dim chk As Long
    On Error GoTo Is1D
    chk = LBound(src, 2)
    On Error GoTo 0

    ' --------------------------------------------------------
    ' Case 4: 2D array (dimension 2 exists)
    ' --------------------------------------------------------
    Dim lb2 As Long, ub2 As Long
    Dim r As Long, c As Long
    lb2 = LBound(src, 2)
    ub2 = UBound(src, 2)

    For r = lb1 To ub1
        For c = lb2 To ub2
            v = src(r, c)
            If VarType(v) <> vbError Then
                If GCD_Try_To_LongLong(v, tmp) Then count = count + 1
            End If
        Next c
    Next r

    GCD_Count_Numeric = count
    Exit Function

Is1D:
    On Error GoTo 0
    For i = lb1 To ub1
        v = src(i)
        If VarType(v) <> vbError Then
            If GCD_Try_To_LongLong(v, tmp) Then count = count + 1
        End If
    Next i
    GCD_Count_Numeric = count
    Exit Function

Try2D:
    On Error GoTo 0
    GCD_Count_Numeric = count

End Function


' ============================================================
' FUNCTION: GCD_Append_Numeric
'
' DESCRIPTION:
' Second-pass collector. Fills a pre-allocated output array
' with valid LongLong-coercible values from src, starting at
' position (offset + 1). The caller must ensure output is large
' enough to hold all values (use GCD_Count_Numeric first).
'
' Accepts the same input shapes as GCD_Count_Numeric:
'   - Range object
'   - Scalar value
'   - 1D array
'   - 2D array
'
' PARAMETERS:
'   src     Input value to scan
'   output  Pre-allocated LongLong array (1-based)
'   offset  Index of the last filled position; updated in place
'
' ============================================================
Private Sub GCD_Append_Numeric( _
    ByVal src As Variant, _
    ByRef output() As LongLong, _
    ByRef offset As Long _
)
    Dim tmp As LongLong
    Dim v As Variant

    ' --------------------------------------------------------
    ' Case 1: Range object
    ' --------------------------------------------------------
    If IsObject(src) Then
        Dim rng As Range, cell As Range
        On Error Resume Next
        Set rng = src
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each cell In rng.Cells
                If Not IsError(cell.Value2) Then
                    If GCD_Try_To_LongLong(cell.Value2, tmp) Then
                        offset = offset + 1
                        output(offset) = tmp
                    End If
                End If
            Next cell
        End If
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' Case 2: Scalar
    ' --------------------------------------------------------
    If Not IsArray(src) Then
        If VarType(src) <> vbError Then
            If GCD_Try_To_LongLong(src, tmp) Then
                offset = offset + 1
                output(offset) = tmp
            End If
        End If
        Exit Sub
    End If

    ' --------------------------------------------------------
    ' Case 3: 1D array
    ' --------------------------------------------------------
    Dim i As Long
    Dim lb1 As Long, ub1 As Long
    On Error GoTo Try2D
    lb1 = LBound(src, 1)
    ub1 = UBound(src, 1)

    Dim chk As Long
    On Error GoTo Is1D
    chk = LBound(src, 2)
    On Error GoTo 0

    ' --------------------------------------------------------
    ' Case 4: 2D array (dimension 2 exists)
    ' --------------------------------------------------------
    Dim lb2 As Long, ub2 As Long
    Dim r As Long, c As Long
    lb2 = LBound(src, 2)
    ub2 = UBound(src, 2)

    For r = lb1 To ub1
        For c = lb2 To ub2
            v = src(r, c)
            If VarType(v) <> vbError Then
                If GCD_Try_To_LongLong(v, tmp) Then
                    offset = offset + 1
                    output(offset) = tmp
                End If
            End If
        Next c
    Next r

    Exit Sub

Is1D:
    On Error GoTo 0
    For i = lb1 To ub1
        v = src(i)
        If VarType(v) <> vbError Then
            If GCD_Try_To_LongLong(v, tmp) Then
                offset = offset + 1
                output(offset) = tmp
            End If
        End If
    Next i
    Exit Sub

Try2D:
    On Error GoTo 0

End Sub


' ============================================================
' FUNCTION: MATH_GCD
'
' DESCRIPTION:
' Public UDF. Computes the GCD of a variable number of integer
' inputs using pairwise Euclidean reduction.
'
' Accepts any combination of:
'   - Scalar integers
'   - Range objects (single cell or multi-cell)
'   - 1D or 2D arrays of integers
'
' Non-numeric, empty and out-of-range values are silently
' ignored. If no valid integers are found, returns #NUM!.
'
' STRATEGY:
' 1. First pass (GCD_Count_Numeric)
'    Count how many valid LongLong values exist across all
'    inputs. This enables exact allocation of the output array
'    without ReDim Preserve inside any loop.
'
' 2. Allocate collected array
'    A single 1-based LongLong array is allocated once.
'
' 3. Second pass (GCD_Append_Numeric)
'    Fill the pre-allocated array with the actual values.
'
' 4. Pairwise GCD reduction
'    Reduce the collected values to a single GCD using GCD_Two.
'    Early exit when result reaches 1 (minimum possible GCD).
'
' RETURN VALUE:
' A LongLong representing the GCD, or CVErr(xlErrNum) if no
' valid inputs were found.
'
' EXAMPLE:
' =MATH_GCD(A1:A10)
' =MATH_GCD(A1:C5, 24, 60, 12)
' =MATH_GCD(18, 24, 42)
'
' VBA usage:
'   Debug.Print MATH_GCD(Range("A1:A10"), 24, 60)
'
' ============================================================
Public Function MATH_GCD(ParamArray values() As Variant) As Variant

    ' --------------------------------------------------------
    ' STEP 1: Count valid values across all inputs
    ' --------------------------------------------------------
    Dim totalCount As Long
    Dim i As Long

    For i = LBound(values) To UBound(values)
        totalCount = totalCount + GCD_Count_Numeric(values(i))
    Next i

    If totalCount = 0 Then
        MATH_GCD = CVErr(xlErrNum)
        Exit Function
    End If

    ' --------------------------------------------------------
    ' STEP 2: Allocate output array exactly once
    ' --------------------------------------------------------
    Dim collected() As LongLong
    ReDim collected(1 To totalCount)

    ' --------------------------------------------------------
    ' STEP 3: Fill collected array
    ' --------------------------------------------------------
    Dim offset As Long
    offset = 0

    For i = LBound(values) To UBound(values)
        GCD_Append_Numeric values(i), collected, offset
    Next i

    ' --------------------------------------------------------
    ' STEP 4: Pairwise GCD reduction
    ' --------------------------------------------------------
    Dim result As LongLong
    result = collected(1)

    For i = 2 To totalCount
        result = GCD_Two(result, collected(i))
        If result = 1 Then Exit For    ' Early exit: 1 is the minimum possible GCD
    Next i

    MATH_GCD = result

End Function


#Else

' ========= Fallback for non-64-bit environments =========
' Please run this in Office 64-bit (Win64).
' This placeholder prevents compile errors on 32-bit Office.
Public Function MATH_GCD(ParamArray values() As Variant) As Variant
    MATH_GCD = CVErr(xlErrNA)
End Function

#End If

' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================
' =======================================================================================================================================








' =======================================================================================================================================
' RADIX CONVERSION ENGINE START
' =======================================================================================================================================


' ============================================================
' FUNCTION: MATH_RADIX_CONVERT
' CATEGORY: Math / Numeral Systems / Base Conversion
'
' DESCRIPTION:
' The MATH_RADIX_CONVERT function converts a number from any
' base (pure or mixed radix) to any other base (pure or mixed).
' It is the primary public interface of the Radix Engine.
'
' A PURE BASE is a single integer >= 2 that applies uniformly
' to every digit position (e.g., base 2, base 10, base 16).
'
' A MIXED BASE is a 1D array where each element defines the
' base of the corresponding digit position. This models systems
' where the base varies by position, such as time
' (seconds [60], minutes [60], hours [24]) or combinatorial
' index decomposition.
'
' SUPPORTED CONVERSIONS:
'   Pure  --> Pure    (e.g., binary --> decimal)
'   Pure  --> Mixed   (e.g., decimal --> [24,60,60])
'   Mixed --> Pure    (e.g., [24,60,60] --> decimal)
'   Mixed --> Mixed
'
' STRATEGY:
' 1. Normalize value --> digits
'    Delegates to Radix_Normalize_Digits, which accepts scalars,
'    Ranges, and arrays and returns a 1-based 1D Long array
'    in MSD --> LSD order.
'
' 2. Normalize bases
'    Delegates to Radix_Normalize_Base for both inputBase and
'    outputBase. Scalars are returned as Long; arrays and
'    ranges are returned as 1-based 1D Long arrays.
'
' 3. Convert to decimal
'    Delegates to MATH_RADIX_TO_DECIMAL, which dispatches to
'    Radix_Pure_To_Decimal or Radix_Mixed_To_Decimal depending
'    on whether baseDef is a scalar or an array.
'
' 4. Convert from decimal to output base
'    Delegates to MATH_RADIX_FROM_DECIMAL, which dispatches to
'    Radix_Decimal_To_Pure or Radix_Decimal_To_Mixed.
'
' PARAMETERS:
'   value       Variant   The number to convert. Accepted forms:
'                           - Scalar integer (e.g., 1010)
'                           - Range containing a scalar or digits
'                           - 1D array of digits (MSD --> LSD)
'
'   inputBase   Variant   The base of the input. Accepted forms:
'                           - Scalar Long >= 2 (pure base)
'                           - 1D array of Long >= 2 (mixed base)
'                           - Range resolving to either of the above
'
'   outputBase  Variant   The target base. Same accepted forms
'                         as inputBase.
'
' RETURN VALUE:
' A 1-based 1D Variant array of digits in MSD --> LSD order
' representing the input value expressed in outputBase.
'
' EXAMPLE:
' ' Binary to decimal:
' MATH_RADIX_CONVERT(Array(1,0,1,0), 2, 10)  ' Returns: (1,0)
'
' ' Decimal to binary:
' MATH_RADIX_CONVERT(10, 10, 2)               ' Returns: (1,0,1,0)
'
' ' Decimal to mixed base [hours=24, minutes=60, seconds=60]:
' MATH_RADIX_CONVERT(3661, 10, Array(24,60,60))
' ' Returns: (1,1,1) --> 1 hour, 1 minute, 1 second
'
' ============================================================

Public Function MATH_RADIX_CONVERT( _
    ByVal value As Variant, _
    ByVal inputBase As Variant, _
    ByVal outputBase As Variant _
) As Variant

    Dim digits As Variant
    Dim decimalValue As Double
    Dim normInputBase As Variant
    Dim normOutputBase As Variant
    
    ' --------------------------------------------------------
    ' STEP 1: Normalize value --> digits
    ' --------------------------------------------------------
 
    digits = Radix_Normalize_Digits(value)
    
    ' --------------------------------------------------------
    ' STEP 2: Normalize bases
    ' --------------------------------------------------------
    
    normInputBase = Radix_Normalize_Base(inputBase)
    normOutputBase = Radix_Normalize_Base(outputBase)
    
    ' --------------------------------------------------------
    ' STEP 3: Convert to decimal
    ' --------------------------------------------------------
    
    decimalValue = MATH_RADIX_TO_DECIMAL(digits, normInputBase)
    
    ' --------------------------------------------------------
    ' STEP 4: Convert to output base
    ' --------------------------------------------------------
    
    MATH_RADIX_CONVERT = MATH_RADIX_FROM_DECIMAL(decimalValue, normOutputBase)

End Function


' ============================================================
' FUNCTION: Radix_Normalize_Base
' CATEGORY: Math / Radix Engine (Internal)
'
' DESCRIPTION:
' The Radix_Normalize_Base function converts any supported base
' representation into one of two canonical forms used
' internally by the Radix Engine:
'
'   - A scalar Long  --> for pure base systems
'   - A 1-based 1D Long array --> for mixed base systems
'
' This normalization shields the rest of the engine from
' input diversity (Range, scalar, 1D array, 2D vector).
'
' PRECONDITIONS (caller must guarantee):
'   baseDef must be scalar, a 1D array, or a Range resolving
'   to either of the above. Rectangular 2D matrices raise
'   ERR_RADIX_BASE_NOT_1D.
'
' STRATEGY:
' 1. Handle Range input
'    Extracts .Value from the Range. If the result is a
'    scalar, returns it as CLng directly. If it is an array,
'    assigns it to arr for further processing.
'
' 2. Handle non-Range input
'    Passes the input through ARR_TO_2D to obtain a uniform
'    2D representation regardless of whether the original
'    input was a scalar, 1D array, or already 2D.
'
' 3. Dispatch by shape
'    Inspects the shape of arr:
'
'      (1 x 1) --> scalar enclosed in 2D: extract and
'                  return as CLng
'      (1 x N) or (N x 1) --> vector: flatten and return
'                              as 1D Long array via
'                              ARR_TO_1D_LONG
'      otherwise --> raise ERR_RADIX_BASE_NOT_1D
'
' RETURN VALUE:
' A Long (pure base) or a 1-based 1D Long array (mixed base).
'
' ============================================================

Private Function Radix_Normalize_Base(ByVal baseDef As Variant) As Variant

    Dim arr As Variant
    Dim shape As Variant

    ' --------------------------------------------------------
    ' CASE 1: Range input
    ' Extract .Value and handle scalar vs. array results
    ' separately, since a single-cell Range returns a scalar
    ' rather than a (1 x 1) array.
    ' --------------------------------------------------------
    If TypeName(baseDef) = "Range" Then
        Dim tmp As Variant
        tmp = baseDef.value
        If Not IsArray(tmp) Then
            ' Single-cell Range: return scalar directly
            Radix_Normalize_Base = CLng(tmp)
            Exit Function
        Else
            ' Multi-cell Range: assign for shape-based dispatch
            arr = tmp
        End If

    Else
        ' --------------------------------------------------------
        ' CASE 2: Non-Range input
        ' Normalize to 2D so all downstream logic operates on
        ' a uniform structure regardless of input shape.
        ' --------------------------------------------------------
        arr = ARR_TO_2D(baseDef)
    End If

    ' --------------------------------------------------------
    ' Inspect shape for dispatch
    ' --------------------------------------------------------
    shape = ARR_SHAPE(arr)

    Dim rows As Long, cols As Long
    rows = shape(1)
    cols = shape(2)

    ' --------------------------------------------------------
    ' CASE 3: Scalar wrapped in a (1 x 1) 2D array
    ' Produced by ARR_TO_2D when the original input was a
    ' scalar (non-Range). Extract and return as Long.
    ' --------------------------------------------------------
    If rows = 1 And cols = 1 Then
        Radix_Normalize_Base = CLng(arr(1, 1))
        Exit Function
    End If

    ' --------------------------------------------------------
    ' CASE 4: Row vector (1 x N) or column vector (N x 1)
    ' Flatten to a 1-based 1D Long array representing a
    ' mixed base definition.
    ' --------------------------------------------------------
    If rows = 1 Or cols = 1 Then
        Radix_Normalize_Base = ARR_TO_1D_LONG(arr)
        Exit Function
    End If

    ' --------------------------------------------------------
    ' CASE 5: Rectangular matrix -- not a valid base
    ' --------------------------------------------------------
    Err.Raise ERR_RADIX_BASE_NOT_1D, , "Base must be scalar or 1D"

End Function


' ============================================================
' FUNCTION: Radix_Normalize_Digits
' CATEGORY: Math / Radix Engine (Internal)
'
' DESCRIPTION:
' The Radix_Normalize_Digits function converts any supported digit
' representation into a canonical 1-based 1D Long array in
' MSD --> LSD order, ready for consumption by the Radix Engine.
'
' Accepted input forms:
'   - Scalar integer  (e.g., 1010  --> digits (1,0,1,0))
'   - Range containing a scalar integer
'   - Range containing a pre-built digit array
'   - 1D or 2D array of digits (passed through as-is)
'
' STRATEGY:
' 1. Resolve input to a digit representation
'    The function first determines the structural type of the
'    input and resolves it to an intermediate digit collection:
'
'      Range (scalar value)  --> TXT_NUM_TO_DIGITS extracts
'                                individual digits from the
'                                integer string representation
'      Range (array value)   --> used directly as digit array
'      Scalar (non-Range)    --> TXT_NUM_TO_DIGITS
'      Array (any shape)     --> used directly as digit array
'
' 2. Flatten to 1D Long array
'    Delegates to ARR_TO_1D_LONG, which handles any
'    remaining structural complexity (scalars wrapped in 2D,
'    row vectors, column vectors) and enforces Long typing
'    via CLng coercion on each element.
'
' RETURN VALUE:
' A 1-based 1D Long array of digits in MSD --> LSD order.
'
' ============================================================

Private Function Radix_Normalize_Digits( _
    ByVal rawDigits As Variant _
) As Variant

    Dim digits As Variant

    ' --------------------------------------------------------
    ' CASE 1: Range input
    ' Extract .Value. A single-cell Range returns a scalar,
    ' which is decomposed into digits via TXT_NUM_TO_DIGITS.
    ' A multi-cell Range is assumed to already contain a
    ' digit array and is used directly.
    ' --------------------------------------------------------
    If TypeName(rawDigits) = "Range" Then
        Dim tmp As Variant
        tmp = rawDigits.value
        If Not IsArray(tmp) Then
            digits = TXT_NUM_TO_DIGITS(tmp)
        Else
            digits = tmp
        End If

    ' --------------------------------------------------------
    ' CASE 2: Scalar non-Range input
    ' Decompose the integer into its individual digits.
    ' --------------------------------------------------------
    ElseIf Not IsArray(rawDigits) Then
        digits = TXT_NUM_TO_DIGITS(rawDigits)

    ' --------------------------------------------------------
    ' CASE 3: Array input
    ' Caller already provides a digit array. Pass through
    ' directly to ARR_TO_1D_LONG for flattening and typing.
    ' --------------------------------------------------------
    Else
        digits = rawDigits
    End If

    ' --------------------------------------------------------
    ' Flatten to a 1-based 1D Long array (MSD --> LSD).
    ' ARR_TO_1D_LONG handles any remaining shape complexity
    ' and enforces Long typing on every element.
    ' --------------------------------------------------------
    Radix_Normalize_Digits = ARR_TO_1D_LONG(digits)

End Function



' ============================================================
' FUNCTION: MATH_RADIX_TO_DECIMAL
' CATEGORY: Math / Radix Engine
'
' DESCRIPTION:
' The MATH_RADIX_TO_DECIMAL function converts a digit array from
' any base (pure or mixed) into its decimal (base 10) integer
' equivalent. It is one of the two public conversion primitives
' of the Radix Engine, alongside MATH_RADIX_FROM_DECIMAL.
'
' PRECONDITION:
' digits must be a 1D array of non-negative integers in
' MSD --> LSD order. Use Radix_Normalize_Digits to produce a
' conformant array from raw user input before calling this
' function directly.
'
' STRATEGY:
' 1. Validate digits
'    If digits is not an array, raises ERR_RADIX_DIGITS_NOT_ARRAY.
'
' 2. Dispatch by base type
'    If baseDef is an array, the base is mixed and the function
'    delegates to Radix_Mixed_To_Decimal.
'    If baseDef is a scalar, the base is pure and the function
'    delegates to Radix_Pure_To_Decimal.
'
' PARAMETERS:
'   digits    Variant   1D array of digits in MSD --> LSD order.
'                       Each digit must be >= 0 and < its
'                       corresponding base value.
'
'   baseDef   Variant   Pure base: a scalar Long >= 2.
'                       Mixed base: a 1D Long array >= 2 per
'                       element, with the same length as digits.
'
' RETURN VALUE:
' A Double representing the decimal value of the input digits.
'
' EXAMPLE:
' MATH_RADIX_TO_DECIMAL(Array(1,0,1,0), 2)   ' Returns: 10
' MATH_RADIX_TO_DECIMAL(Array(1,1,1), Array(24,60,60))
' ' Returns: 3661  (1 hour + 1 minute + 1 second in seconds)
'
' ============================================================

Public Function MATH_RADIX_TO_DECIMAL( _
    ByVal digits As Variant, _
    ByVal baseDef As Variant _
) As Double

    ' --------------------------------------------------------
    ' Validate: digits must be an array
    ' Scalars are not accepted here -- use Radix_Normalize_Digits
    ' to convert a scalar integer to a digit array first.
    ' --------------------------------------------------------
    If Not IsArray(digits) Then
        Err.Raise ERR_RADIX_DIGITS_NOT_ARRAY, , "Digits must be an array"
    End If

    ' --------------------------------------------------------
    ' Dispatch by base type
    ' --------------------------------------------------------
    If IsArray(baseDef) Then
        ' Mixed base: each digit position has its own base
        MATH_RADIX_TO_DECIMAL = Radix_Mixed_To_Decimal(digits, baseDef)
    Else
        ' Pure base: single base value applies to all positions
        MATH_RADIX_TO_DECIMAL = Radix_Pure_To_Decimal(digits, CLng(baseDef))
    End If

End Function



' ============================================================
' FUNCTION: Radix_Pure_To_Decimal
' CATEGORY: Math / Radix Engine (Internal)
'
' DESCRIPTION:
' The Radix_Pure_To_Decimal function converts a digit array
' from a pure base system into its decimal equivalent using
' iterative accumulation (Horner's method):
'
'   result = result * base + digit
'
' This is equivalent to evaluating the polynomial:
'
'   d[0]*base^(n-1) + d[1]*base^(n-2) + ... + d[n-1]*base^0
'
' but requires only N multiplications instead of computing
' each power separately.
'
' PRECONDITIONS (caller must guarantee):
'   digits must be a 1D array in MSD --> LSD order.
'   baseVal must be >= 2 (validated internally).
'   Each digit must satisfy 0 <= digit < baseVal.
'
' STRATEGY:
' 1. Validate base
'    Raises ERR_RADIX_BASE_TOO_SMALL if baseVal < 2.
'
' 2. Iterate MSD --> LSD
'    For each digit, validate it is within [0, baseVal),
'    then accumulate using:
'
'      result = result * baseVal + digit
'
' PARAMETERS:
'   digits    Variant   1D array of digits in MSD --> LSD order.
'   baseVal   Long      The base of the input numeral system.
'                       Must be >= 2.
'
' RETURN VALUE:
' A Double representing the decimal value of the input digits.
'
' EXAMPLE:
' Radix_Pure_To_Decimal(Array(1,0,1,0), 2)   ' Returns: 10
' Radix_Pure_To_Decimal(Array(2,5,5), 16)    ' Returns: 597
'
' ============================================================
Private Function Radix_Pure_To_Decimal( _
    ByVal digits As Variant, _
    ByVal baseVal As Long _
) As Double

    ' --------------------------------------------------------
    ' Validate base
    ' --------------------------------------------------------
    If baseVal < 2 Then
        Err.Raise ERR_RADIX_BASE_TOO_SMALL, , "Base must be >= 2"
    End If

    Dim i As Long
    Dim result As Double
    Dim d As Long

    Dim lb As Long, ub As Long
    lb = LBound(digits)
    ub = UBound(digits)

    result = 0

    ' --------------------------------------------------------
    ' Accumulate digits MSD --> LSD using Horner's method:
    ' result = result * base + digit
    ' --------------------------------------------------------
    For i = lb To ub

        d = CLng(digits(i))

        ' Validate digit is within the valid range for this base
        If d < 0 Or d >= baseVal Then
            Err.Raise ERR_RADIX_DIGIT_OUT_OF_RANGE, , "Digit out of range for base"
        End If

        result = result * baseVal + d

    Next i

    Radix_Pure_To_Decimal = result

End Function



' ============================================================
' FUNCTION: Radix_Mixed_To_Decimal
' CATEGORY: Math / Radix Engine (Internal)
'
' DESCRIPTION:
' The Radix_Mixed_To_Decimal function converts a digit array
' from a mixed radix system into its decimal equivalent using
' right-to-left accumulation of positional weights.
'
' In a mixed radix system each digit position i has its own
' base b[i]. The decimal value is computed as:
'
'   result = d[n-1]*1 + d[n-2]*b[n-1] + d[n-3]*b[n-1]*b[n-2] + ...
'
' where the positional weight of each digit is the product
' of all bases to its right. The traversal starts at the
' least significant position (rightmost) and accumulates
' the weight (divisor) leftward.
'
' EXAMPLE (time: hours, minutes, seconds):
'   digits  = (1, 1, 1)
'   baseArr = (24, 60, 60)
'
'   Position 2 (seconds): 1 * 1        =    1
'   Position 1 (minutes): 1 * 60       =   60
'   Position 0 (hours):   1 * 60 * 60  = 3600
'                                        ----
'                                        3661 seconds total
'
' PRECONDITIONS (caller must guarantee):
'   digits and baseArr must have the same length.
'   digits and baseArr must be 1D arrays.
'   Both may use any lower bound (index-safe access used).
'
' STRATEGY:
' 1. Validate length consistency
'    Raises ERR_RADIX_SIZE_MISMATCH if lengths differ.
'
' 2. Traverse right to left (LSD --> MSD)
'    At each position:
'      a. Compute index-safe access into digits and baseArr
'         since their lower bounds may differ.
'      b. Validate that b >= 2 and 0 <= d < b.
'      c. Accumulate: result = result + d * divisor
'      d. Update positional weight: divisor = divisor * b
'
' PARAMETERS:
'   digits    Variant   1D array of digits in MSD --> LSD order.
'                       Lower bound may be anything.
'   baseArr   Variant   1D array of base values, one per digit
'                       position. Each value must be >= 2.
'                       Lower bound may be anything.
'
' RETURN VALUE:
' A Double representing the decimal value of the input digits.
'
' ============================================================
Private Function Radix_Mixed_To_Decimal( _
    ByVal digits As Variant, _
    ByVal baseArr As Variant _
) As Double

    Dim i As Long
    Dim result As Double
    Dim divisor As Double
    Dim d As Long
    Dim b As Long

    Dim lb As Long, ub As Long
    lb = LBound(baseArr)
    ub = UBound(baseArr)

    Dim d_lb As Long, d_ub As Long
    d_lb = LBound(digits)
    d_ub = UBound(digits)

    ' --------------------------------------------------------
    ' Validate: digits and baseArr must have the same length
    ' --------------------------------------------------------
    If (d_ub - d_lb + 1) <> (ub - lb + 1) Then
        Err.Raise ERR_RADIX_SIZE_MISMATCH, , "Digits and base length mismatch"
    End If

    result = 0
    divisor = 1

    ' --------------------------------------------------------
    ' Traverse LSD --> MSD accumulating positional weights.
    ' divisor holds the weight of the current position:
    '   starts at 1 (units), multiplied by each base as we
    '   move left, so position k has weight b[k+1]*b[k+2]*...
    ' --------------------------------------------------------
    For i = ub To lb Step -1

        ' Index-safe access: digits and baseArr may have
        ' different lower bounds so offsets are computed
        ' explicitly from their respective lb values.
        d = CLng(digits(d_lb + (i - lb)))
        b = CLng(baseArr(i))

        ' Validate base for this position
        If b < 2 Then
            Err.Raise ERR_RADIX_MIXED_BASE_INVALID, , "Base values must be >= 2"
        End If

        ' Validate digit is within range for this position's base
        If d < 0 Or d >= b Then
            Err.Raise ERR_RADIX_MIXED_DIGIT_INVALID, , "Digit out of range for mixed base"
        End If

        result = result + d * divisor
        divisor = divisor * b

    Next i

    Radix_Mixed_To_Decimal = result

End Function




' ============================================================
' FUNCTION: MATH_RADIX_FROM_DECIMAL
' CATEGORY: Math / Radix Engine
'
' DESCRIPTION:
' The MATH_RADIX_FROM_DECIMAL function converts a non-negative
' decimal integer into its representation in any target base
' (pure or mixed). It is one of the two public conversion
' primitives of the Radix Engine, alongside MATH_RADIX_TO_DECIMAL.
'
' PRECONDITION:
' decimalValue must be a non-negative integer. Fractional
' values raise ERR_RADIX_NON_INTEGER (enforced internally
' by Radix_Decimal_To_Pure). Negative values raise
' ERR_RADIX_NEGATIVE_VALUE (validated here).
'
' STRATEGY:
' 1. Validate sign
'    Raises ERR_RADIX_NEGATIVE_VALUE if decimalValue < 0.
'
' 2. Dispatch by base type
'    If baseDef is an array, the target base is mixed and
'    the function delegates to Radix_Decimal_To_Mixed.
'    If baseDef is a scalar, validates it is >= 2 and
'    delegates to Radix_Decimal_To_Pure.
'
' PARAMETERS:
'   decimalValue  Double    The non-negative integer to convert.
'
'   baseDef       Variant   Pure base: a scalar Long >= 2.
'                           Mixed base: a 1D Long array >= 2
'                           per element.
'
' RETURN VALUE:
' A 1-based 1D Variant array of digits in MSD --> LSD order
' representing decimalValue in the target base.
'
' EXAMPLE:
' MATH_RADIX_FROM_DECIMAL(10, 2)
' ' Returns: (1,0,1,0)
'
' MATH_RADIX_FROM_DECIMAL(3661, Array(24,60,60))
' ' Returns: (1,1,1)  --> 1 hour, 1 minute, 1 second
'
' ============================================================
Public Function MATH_RADIX_FROM_DECIMAL( _
    ByVal decimalValue As Double, _
    ByVal baseDef As Variant _
) As Variant

    ' --------------------------------------------------------
    ' Validate sign: negative values are not supported
    ' --------------------------------------------------------
    If decimalValue < 0 Then
        Err.Raise ERR_RADIX_NEGATIVE_VALUE, , "Negative values not supported"
    End If

    ' --------------------------------------------------------
    ' Dispatch by base type
    ' --------------------------------------------------------
    If IsArray(baseDef) Then
        ' Mixed base: each digit position has its own base
        MATH_RADIX_FROM_DECIMAL = Radix_Decimal_To_Mixed(decimalValue, baseDef)
    Else
        ' Pure base: validate scalar and delegate
        If CLng(baseDef) < 2 Then
            Err.Raise ERR_RADIX_BASE_TOO_SMALL, , "Base must be >= 2"
        End If
        MATH_RADIX_FROM_DECIMAL = Radix_Decimal_To_Pure(decimalValue, CLng(baseDef))
    End If

End Function



' ============================================================
' FUNCTION: Radix_Decimal_To_Pure
' CATEGORY: Math / Radix Engine (Internal)
'
' DESCRIPTION:
' The Radix_Decimal_To_Pure function converts a non-negative
' decimal integer into its representation in a pure base
' using repeated division.
'
' The algorithm extracts digits from least significant to
' most significant (LSD --> MSD) by repeatedly dividing by
' the base and collecting the remainders. The digits are
' collected into a fixed-size buffer and then read back in
' reverse order to produce the final MSD --> LSD result.
'
' PRECONDITIONS (caller must guarantee):
'   decimalValue must be >= 0 (validated by MATH_RADIX_FROM_DECIMAL).
'   baseVal must be >= 2 (validated by MATH_RADIX_FROM_DECIMAL).
'   decimalValue must be an integer (validated internally).
'
' STRATEGY:
' 1. Validate integer input
'    Raises ERR_RADIX_NON_INTEGER if decimalValue has a
'    fractional part (decimalValue <> Fix(decimalValue)).
'
' 2. Handle zero
'    Returns (0) immediately as a special case, since the
'    division loop would produce no iterations for value = 0.
'
' 3. Extract digits into buffer (LSD --> MSD)
'    Repeatedly computes value Mod baseVal to extract the
'    current LSD, then divides value by baseVal to shift
'    right. Digits are stored right-to-left in a fixed
'    buffer(1 To 64), using pos as a descending write cursor.
'    Raises ERR_RADIX_OVERFLOW if pos reaches 0, meaning
'    the result requires more than 64 digits.
'
' 4. Copy buffer to result (MSD --> LSD)
'    Reads the filled portion of the buffer left-to-right
'    into a 1-based result array of exact size.
'
' PARAMETERS:
'   decimalValue  Double   Non-negative integer to convert.
'   baseVal       Long     Target base. Must be >= 2.
'
' RETURN VALUE:
' A 1-based 1D Variant array of digits in MSD --> LSD order.
'
' EXAMPLE:
' Radix_Decimal_To_Pure(10, 2)    ' Returns: (1,0,1,0)
' Radix_Decimal_To_Pure(255, 16)  ' Returns: (15,15)
' Radix_Decimal_To_Pure(0, 10)    ' Returns: (0)
'
' ============================================================
Private Function Radix_Decimal_To_Pure( _
    ByVal decimalValue As Double, _
    ByVal baseVal As Long _
) As Variant

    ' --------------------------------------------------------
    ' Validate: decimalValue must be an integer
    ' --------------------------------------------------------
    If decimalValue <> Fix(decimalValue) Then
        Err.Raise ERR_RADIX_NON_INTEGER, , "Decimal value must be integer"
    End If

    Dim result() As Variant
    Dim value As Double
    value = decimalValue

    ' --------------------------------------------------------
    ' Special case: zero produces a single digit (0)
    ' The division loop below would produce no iterations,
    ' so this case must be handled explicitly.
    ' --------------------------------------------------------
    If value = 0 Then
        ReDim result(1 To 1)
        result(1) = 0
        Radix_Decimal_To_Pure = result
        Exit Function
    End If

    ' --------------------------------------------------------
    ' Extract digits LSD --> MSD into a fixed-size buffer.
    ' pos is a descending write cursor starting at 64.
    ' Each iteration writes one digit and moves pos left.
    ' --------------------------------------------------------
    Dim buffer(1 To 64) As Long
    Dim pos As Long: pos = 64

    Do While value > 0
        ' Guard against buffer overflow before writing
        If pos < 1 Then
            Err.Raise ERR_RADIX_OVERFLOW, , _
                "Radix_Decimal_To_Pure: result exceeds 64-digit buffer " & _
                "(value=" & decimalValue & ", base=" & baseVal & ")"
        End If
        buffer(pos) = value Mod baseVal
        value = Int(value / baseVal)
        pos = pos - 1
    Loop

    ' --------------------------------------------------------
    ' Copy filled buffer portion into result (MSD --> LSD).
    ' The digits occupy buffer(pos+1) to buffer(64).
    ' --------------------------------------------------------
    Dim digitCount As Long
    digitCount = 64 - pos

    ReDim result(1 To digitCount)
    Dim i As Long
    For i = 1 To digitCount
        result(i) = buffer(pos + i)
    Next i

    Radix_Decimal_To_Pure = result

End Function



' ============================================================
' FUNCTION: Radix_Decimal_To_Mixed
' CATEGORY: Math / Radix Engine (Internal)
'
' DESCRIPTION:
' The Radix_Decimal_To_Mixed function converts a non-negative
' decimal integer into its mixed radix representation using
' a pre-computed divisor system.
'
' For a mixed base array [b0, b1, b2], the positional weight
' of each digit is the product of all bases to its right:
'
'   divisor[0] = b1 * b2
'   divisor[1] = b2
'   divisor[2] = 1
'
' Each digit is then extracted as:
'
'   digit[i] = (decimalValue \ divisor[i]) Mod base[i]
'
' This is the same decomposition used in Cartesian product
' index mapping, where each dimension plays the role of a
' base and the row index plays the role of the decimal value.
'
' PRECONDITIONS (caller must guarantee):
'   decimalValue must be >= 0.
'   baseArr must be a 1D array of Long values, each >= 2.
'   decimalValue must be representable within the capacity
'   of the mixed base system (i.e., < product of all bases).
'   No overflow validation is performed here.
'
' STRATEGY:
' 1. Build divisors (right to left)
'    divisors(ub) = 1. Each position i gets the product of
'    all bases to its right: divisors(i) = divisors(i+1) * base(i+1).
'
' 2. Extract digits (left to right)
'    For each position i, compute:
'      digit = (decimalValue \ divisors(i)) Mod base(i)
'    Store in result using a 1-based offset since baseArr
'    may have any lower bound.
'
' PARAMETERS:
'   decimalValue  Double    Non-negative integer to convert.
'   baseArr       Variant   1D array of base values (Long >= 2),
'                           one per output digit position.
'                           May use any lower bound.
'
' RETURN VALUE:
' A 1-based 1D Variant array of digits in MSD --> LSD order.
'
' EXAMPLE:
' Radix_Decimal_To_Mixed(3661, Array(24,60,60))
' ' divisors: (3600, 60, 1)
' ' digits:   (1, 1, 1)  --> 1 hour, 1 minute, 1 second
'
' ============================================================
Private Function Radix_Decimal_To_Mixed( _
    ByVal decimalValue As Double, _
    ByVal baseArr As Variant _
) As Variant

    Dim lb As Long, ub As Long
    lb = LBound(baseArr)
    ub = UBound(baseArr)

    Dim numDims As Long
    numDims = ub - lb + 1

    Dim result() As Variant
    ReDim result(1 To numDims)

    ' divisors(i) holds the positional weight of digit i:
    ' the product of all bases to the right of position i.
    Dim divisors() As Double
    ReDim divisors(lb To ub)

    Dim i As Long

    ' --------------------------------------------------------
    ' STEP 1: Build positional weights right to left.
    ' The rightmost position always has weight 1 (units).
    ' Each position to the left is base(i+1) times larger.
    ' --------------------------------------------------------
    divisors(ub) = 1

    For i = ub - 1 To lb Step -1
        divisors(i) = divisors(i + 1) * CLng(baseArr(i + 1))
    Next i

    ' --------------------------------------------------------
    ' STEP 2: Extract each digit using its positional weight.
    ' Integer-divide to isolate the contribution of position i,
    ' then Mod to keep only the digit for that position.
    ' The offset (i - lb + 1) maps baseArr's lower bound to
    ' the 1-based result array.
    ' --------------------------------------------------------
    For i = lb To ub
        result(i - lb + 1) = (decimalValue \ divisors(i)) Mod CLng(baseArr(i))
    Next i

    Radix_Decimal_To_Mixed = result

End Function

' =======================================================================================================================================
' RADIX CONVERSION ENGINE END
' =======================================================================================================================================


