Attribute VB_Name = "TEXT"
Option Explicit

' ============================================================
' MODULE: TEXT
' CATEGORY: Text / String Utilities
'
' PURPOSE:
' This module provides utility functions for extracting and
' transforming subsets of characters from strings. It also
' serves as the numeric-to-digits bridge used by the Radix
' Engine in the MATH module.
'
' ------------------------------------------------------------
' PUBLIC API
' ------------------------------------------------------------
'
'   TXT_STRIP_DIGITS
'     Extracts all non-numeric characters from a string.
'     Preserves letters, spaces, punctuation and symbols.
'
'   TXT_EXTRACT_DIGITS
'     Extracts all digit characters (0-9) from a string and
'     returns them concatenated. Keeps all digits regardless
'     of position. For the first sequence only, use
'     TXT_EXTRACT_FIRSTNUM.
'
'   TXT_EXTRACT_ALPHA
'     Extracts all ASCII alphabetic characters (A-Z, a-z).
'     Discards digits, spaces, punctuation and non-ASCII
'     letters (e.g., é, ñ, ü).
'
'   TXT_EXTRACT_ALPHANUM
'     Extracts all ASCII alphanumeric characters (A-Z, a-z,
'     0-9). Discards spaces, punctuation and non-ASCII
'     characters.
'
'   TXT_EXTRACT_FIRSTNUM
'     Extracts the first contiguous sequence of digits (0-9)
'     from a string. Stops scanning as soon as the sequence
'     ends.
'
'   TXT_NUM_TO_DIGITS
'     Converts a non-negative integer into a 1-based 1D
'     array of its individual digits in MSD --> LSD order.
'     Used internally by Normalize_Digits in MATH.
'
' ------------------------------------------------------------
' DESIGN CONVENTIONS
' ------------------------------------------------------------
'
' 1. DIGIT DETECTION
'    All functions use range comparison (ch >= "0" And
'    ch <= "9") or Like "[A-Za-z0-9]" for character
'    classification. IsNumeric is deliberately avoided
'    because it accepts non-digit characters such as
'    ".", "+", "-" and currency symbols.
'
' 2. ASCII ONLY
'    Character classification covers ASCII only. Accented
'    or non-Latin characters (e.g., é, ñ, ü) are treated
'    as non-alphabetic by all functions in this module.
'
' 3. RETURN TYPES
'    TXT_STRIP_DIGITS, TXT_EXTRACT_DIGITS, TXT_EXTRACT_ALPHA, TXT_EXTRACT_ALPHANUM
'    and TXT_EXTRACT_FIRSTNUM return String.
'    TXT_NUM_TO_DIGITS returns a 1-based 1D Variant array.
'
' 4. ERROR SIGNALING
'    Only TXT_NUM_TO_DIGITS raises an error (ERR_TXT_INVALID_NUMERIC)
'    when the input contains non-digit characters after sign
'    stripping. All other functions silently discard
'    non-matching characters.
'
' ------------------------------------------------------------
' DEPENDENCIES
' ------------------------------------------------------------
'
'   ERRORS module:
'     ERR_TXT_INVALID_NUMERIC
'
' ------------------------------------------------------------
' RELATED MODULES
' ------------------------------------------------------------
'
'   MATH module:
'     Normalize_Digits calls TXT_NUM_TO_DIGITS to convert
'     scalar and Range inputs into digit arrays before
'     passing them to the Radix Engine.
'
' ============================================================




' ===========================================================================================
' ===========================================================================================
' ===========================================================================================





' ============================================================
' FUNCTION: TXT_STRIP_DIGITS
' CATEGORY: Text / Extraction
'
' DESCRIPTION:
' The TXT_STRIP_DIGITS function extracts all non-numeric characters
' from a string and returns them as a new string, preserving
' their original order and any spacing.
'
' Numeric characters (0-9) are the only characters removed.
' All other characters -- letters, spaces, punctuation, and
' special symbols -- are preserved.
'
' STRATEGY:
' Iterates over each character in the input string. Any
' character that is not a digit (0-9) is appended to the
' result. Digit detection uses a range comparison
' (ch < "0" Or ch > "9") rather than IsNumeric, which would
' incorrectly classify characters like ".", "+", "-" as
' numeric.
'
' PARAMETERS:
'   str   String   The input string to process.
'
' RETURN VALUE:
' A String containing all non-numeric characters from str,
' in their original order. Returns an empty string if str
' contains only digits.
'
' EXAMPLE:
' TXT_STRIP_DIGITS("hello 123 world")  ' Returns: "hello  world"
' TXT_STRIP_DIGITS("price: $50.00")    ' Returns: "price: $."
' TXT_STRIP_DIGITS("abc")              ' Returns: "abc"
' TXT_STRIP_DIGITS("12345")            ' Returns: ""
'
' ============================================================

Public Function TXT_STRIP_DIGITS(ByVal str As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String

    ' --------------------------------------------------------
    ' Iterate over each character and keep non-digit ones.
    ' Range comparison (ch < "0" Or ch > "9") is used instead
    ' of IsNumeric, which accepts ".", "+", "-" as numeric.
    ' --------------------------------------------------------
    For i = 1 To Len(str)

        ch = Mid$(str, i, 1)

        If ch < "0" Or ch > "9" Then
            result = result & ch
        End If

    Next i

    TXT_STRIP_DIGITS = result

End Function


' ============================================================
' FUNCTION: TXT_EXTRACT_DIGITS
' CATEGORY: Text / Extraction
'
' DESCRIPTION:
' The TXT_EXTRACT_DIGITS function extracts all numeric characters
' (0-9) from a string and returns them concatenated as a
' new string, preserving their original order.
'
' Non-numeric characters are discarded. All digit characters
' are kept regardless of their position or surrounding
' context.
'
' Note that this function extracts ALL digits across the
' entire string. To extract only the first contiguous
' numeric sequence, use TXT_EXTRACT_FIRSTNUM instead.
'
' STRATEGY:
' Iterates over each character in the input string. Any
' character that is a digit (0-9) is appended to the result.
' Digit detection uses a range comparison
' (ch >= "0" And ch <= "9") rather than IsNumeric, which
' would incorrectly classify characters like ".", "+", "-"
' as numeric.
'
' PARAMETERS:
'   str   String   The input string to process.
'
' RETURN VALUE:
' A String containing all digit characters from str in their
' original order. Returns an empty string if str contains
' no digits.
'
' EXAMPLE:
' TXT_EXTRACT_DIGITS("abc123def456")   ' Returns: "123456"
' TXT_EXTRACT_DIGITS("price: $50.00")  ' Returns: "5000"
' TXT_EXTRACT_DIGITS("hello")          ' Returns: ""
' TXT_EXTRACT_DIGITS("12345")          ' Returns: "12345"
'
' ============================================================

Public Function TXT_EXTRACT_DIGITS(ByVal str As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String

    ' --------------------------------------------------------
    ' Iterate over each character and keep digit ones.
    ' Range comparison (ch >= "0" And ch <= "9") is used
    ' instead of IsNumeric, which accepts ".", "+", "-"
    ' as numeric.
    ' --------------------------------------------------------
    For i = 1 To Len(str)

        ch = Mid$(str, i, 1)

        If ch >= "0" And ch <= "9" Then
            result = result & ch
        End If

    Next i

    TXT_EXTRACT_DIGITS = result

End Function


' ============================================================
' FUNCTION: TXT_EXTRACT_ALPHA
' CATEGORY: Text / Extraction
'
' DESCRIPTION:
' The TXT_EXTRACT_ALPHA function extracts all alphabetic characters
' (A-Z and a-z) from a string and returns them concatenated
' as a new string, preserving their original order and case.
'
' All non-alphabetic characters -- digits, spaces, punctuation,
' and special symbols -- are discarded.
'
' Note that this function only recognizes ASCII letters (A-Z,
' a-z). Accented or non-Latin characters (e.g., é, ñ, ü)
' are not recognized as alphabetic and will be discarded.
'
' STRATEGY:
' Iterates over each character in the input string. Any
' character matching the pattern [A-Za-z] is appended to
' the result. Detection uses the Like operator with a
' character class pattern, which is more readable and
' explicit than ASCII range comparisons for letter detection.
'
' PARAMETERS:
'   str   String   The input string to process.
'
' RETURN VALUE:
' A String containing all ASCII alphabetic characters from
' str in their original order. Returns an empty string if
' str contains no alphabetic characters.
'
' EXAMPLE:
' TXT_EXTRACT_ALPHA("hello 123 world")  ' Returns: "helloworld"
' TXT_EXTRACT_ALPHA("price: $50.00")    ' Returns: "price"
' TXT_EXTRACT_ALPHA("abc")              ' Returns: "abc"
' TXT_EXTRACT_ALPHA("12345")            ' Returns: ""
'
' ============================================================

Public Function TXT_EXTRACT_ALPHA(ByVal str As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String

    ' --------------------------------------------------------
    ' Iterate over each character and keep alphabetic ones.
    ' Like "[A-Za-z]" matches only ASCII letters -- accented
    ' or non-Latin characters are not matched and are
    ' discarded.
    ' --------------------------------------------------------
    For i = 1 To Len(str)

        ch = Mid$(str, i, 1)

        If ch Like "[A-Za-z]" Then
            result = result & ch
        End If

    Next i

    TXT_EXTRACT_ALPHA = result

End Function


' ============================================================
' FUNCTION: TXT_EXTRACT_ALPHANUM
' CATEGORY: Text / Extraction
'
' DESCRIPTION:
' The TXT_EXTRACT_ALPHANUM function extracts all alphanumeric
' characters (A-Z, a-z, and 0-9) from a string and returns
' them concatenated as a new string, preserving their
' original order and case.
'
' All non-alphanumeric characters -- spaces, punctuation,
' and special symbols -- are discarded.
'
' Note that this function only recognizes ASCII letters (A-Z,
' a-z) and ASCII digits (0-9). Accented or non-Latin
' characters (e.g., é, ñ, ü) are not recognized as
' alphabetic and will be discarded.
'
' STRATEGY:
' Iterates over each character in the input string. Any
' character matching the pattern [A-Za-z0-9] is appended
' to the result. Detection uses the Like operator with a
' character class pattern, consistent with TXT_EXTRACT_ALPHA.
'
' PARAMETERS:
'   str   String   The input string to process.
'
' RETURN VALUE:
' A String containing all ASCII alphanumeric characters from
' str in their original order. Returns an empty string if
' str contains no alphanumeric characters.
'
' EXAMPLE:
' TXT_EXTRACT_ALPHANUM("hello 123 world")  ' Returns: "hello123world"
' TXT_EXTRACT_ALPHANUM("price: $50.00")    ' Returns: "price5000"
' TXT_EXTRACT_ALPHANUM("abc")              ' Returns: "abc"
' TXT_EXTRACT_ALPHANUM("!@#$%")            ' Returns: ""
'
' ============================================================

Public Function TXT_EXTRACT_ALPHANUM(ByVal str As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String

    ' --------------------------------------------------------
    ' Iterate over each character and keep alphanumeric ones.
    ' Like "[A-Za-z0-9]" matches ASCII letters and digits.
    ' Accented or non-Latin characters are not matched and
    ' are discarded.
    ' --------------------------------------------------------
    For i = 1 To Len(str)

        ch = Mid$(str, i, 1)

        If ch Like "[A-Za-z0-9]" Then
            result = result & ch
        End If

    Next i

    TXT_EXTRACT_ALPHANUM = result

End Function


' ============================================================
' FUNCTION: TXT_EXTRACT_FIRSTNUM
' CATEGORY: Text / Extraction
'
' DESCRIPTION:
' The TXT_EXTRACT_FIRSTNUM function extracts the first contiguous
' sequence of digit characters (0-9) from a string and
' returns it as a string.
'
' Only the first numeric sequence is returned. Scanning stops
' immediately after the first sequence ends, so any digits
' appearing later in the string are ignored.
'
' Note that this function returns a String, not a numeric
' type. To convert the result to a number, use CLng or CDbl
' on the return value.
'
' STRATEGY:
' Iterates over each character in the input string using a
' started flag to track whether a numeric sequence has begun:
'
'   - If the current character is a digit, append it to the
'     result and set started to True.
'   - If started is True and the current character is not a
'     digit, the first sequence has ended -- exit the loop.
'   - If started is False and the current character is not a
'     digit, continue scanning.
'
' Digit detection uses a range comparison
' (ch >= "0" And ch <= "9") rather than IsNumeric, which
' would incorrectly classify characters like ".", "+", "-"
' as numeric and produce unexpected results for inputs like
' "abc+123".
'
' PARAMETERS:
'   str   String   The input string to process.
'
' RETURN VALUE:
' A String containing the first contiguous digit sequence
' found in str. Returns an empty string if str contains
' no digits.
'
' EXAMPLE:
' TXT_EXTRACT_FIRSTNUM("abc123def456")  ' Returns: "123"
' TXT_EXTRACT_FIRSTNUM("99 bottles")    ' Returns: "99"
' TXT_EXTRACT_FIRSTNUM("price $50.00")  ' Returns: "50"
' TXT_EXTRACT_FIRSTNUM("hello")         ' Returns: ""
'
' ============================================================

Public Function TXT_EXTRACT_FIRSTNUM(ByVal str As String) As String

    Dim result As String
    Dim i As Long
    Dim ch As String
    Dim started As Boolean

    ' --------------------------------------------------------
    ' Iterate over each character tracking whether a numeric
    ' sequence has started. Exit as soon as the first
    ' sequence ends.
    ' --------------------------------------------------------
    For i = 1 To Len(str)

        ch = Mid$(str, i, 1)

        If ch >= "0" And ch <= "9" Then

            ' Current character is a digit -- accumulate it
            result = result & ch
            started = True

        ElseIf started Then
            ' First numeric sequence has ended -- stop scanning
            
            ' Stop if the number sequence has ended
            ' --- THE "STOP" RULE ---
            ' If we reach this point, it means the current character is NOT a number.
            ' We check the "started" flag:
            ' 1. If 'started' is True, it means we WERE reading numbers and just hit a letter/space.
            ' 2. This signifies the end of the first numeric sequence.
            ' 3. We use "Exit For" to stop the loop immediately and keep only the first number found.
            Exit For

        End If

    Next i

    TXT_EXTRACT_FIRSTNUM = result

End Function



' ============================================================
' FUNCTION: TXT_NUM_TO_DIGITS
' CATEGORY: Text / Conversion
'
' DESCRIPTION:
' The TXT_NUM_TO_DIGITS function converts a non-negative
' integer into a 1-based 1D array of its individual digits
' (0-9), one element per digit position in MSD --> LSD order.
'
' This function is the primary bridge between numeric input
' and the Radix Engine. Normalize_Digits in MATH calls this
' function whenever a scalar or single-cell Range is provided
' as the value to convert.
'
' Negative values are accepted -- the sign is stripped before
' processing and the absolute digits are returned.
'
' STRATEGY:
' 1. Normalize Range input
'    If the input is a Range, extracts .Value before
'    processing. A single-cell Range returns a scalar.
'
' 2. Convert to string and trim
'    Converts the value to a trimmed string via CStr. Returns
'    Empty immediately if the result is empty.
'
' 3. Strip leading minus sign
'    If the string starts with "-", removes it. The sign is
'    not represented in the output array.
'
' 4. Allocate result array
'    Creates a 1-based array with one slot per character.
'
' 5. Parse characters --> digits
'    Iterates over each character. Non-digit characters
'    raise ERR_TXT_INVALID_NUMERIC. Valid digits are
'    converted to Long via CLng and stored.
'
' PARAMETERS:
'   value   Variant   The integer to decompose. Accepted forms:
'                       - Scalar integer (e.g., 1010)
'                       - Negative integer (sign is stripped)
'                       - Range containing a scalar integer
'
' RETURN VALUE:
' A 1-based 1D Variant array of Long values in MSD --> LSD
' order. Each element is a single digit (0-9).
' Returns Empty if value resolves to an empty string.
' Raises ERR_TXT_INVALID_NUMERIC if any non-digit character
' is found after sign stripping (e.g., decimals, letters).
'
' EXAMPLE:
' TXT_NUM_TO_DIGITS(1010)   ' Returns: (1, 0, 1, 0)
' TXT_NUM_TO_DIGITS(-42)    ' Returns: (4, 2)
' TXT_NUM_TO_DIGITS(0)      ' Returns: (0)
'
' ============================================================

Public Function TXT_NUM_TO_DIGITS( _
    ByVal value As Variant _
) As Variant

    Dim strValue As String
    Dim i As Long
    Dim result() As Variant
    Dim ch As String

    ' --------------------------------------------------------
    ' STEP 1: Normalize Range input
    ' A single-cell Range returns a scalar via .Value.
    ' --------------------------------------------------------
    If TypeName(value) = "Range" Then
        Dim rangeVal As Variant
        rangeVal = value.value
        value = rangeVal
    End If

    ' --------------------------------------------------------
    ' STEP 2: Convert to string and trim
    ' --------------------------------------------------------
    strValue = Trim$(CStr(value))

    If Len(strValue) = 0 Then Exit Function

    ' --------------------------------------------------------
    ' STEP 3: Strip leading minus sign
    ' The sign is not represented in the output array.
    ' --------------------------------------------------------
    If Left$(strValue, 1) = "-" Then
        strValue = Mid$(strValue, 2)
    End If

    ' --------------------------------------------------------
    ' STEP 4: Allocate result array (one slot per digit)
    ' --------------------------------------------------------
    ReDim result(1 To Len(strValue))

    ' --------------------------------------------------------
    ' STEP 5: Parse characters --> digits
    ' Any non-digit character raises ERR_TXT_INVALID_NUMERIC.
    ' This rejects decimals, spaces, and any other non-integer
    ' input that survived the sign-stripping step.
    ' --------------------------------------------------------
    For i = 1 To Len(strValue)

        ch = Mid$(strValue, i, 1)

        If ch < "0" Or ch > "9" Then
            Err.Raise ERR_TXT_INVALID_NUMERIC, , _
                "TXT_NUM_TO_DIGITS: non-digit character """ & ch & _
                """ at position " & i & " in """ & strValue & """"
        End If

        result(i) = CLng(ch)

    Next i

    TXT_NUM_TO_DIGITS = result

End Function

