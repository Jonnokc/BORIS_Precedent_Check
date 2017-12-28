Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DESCRIPTION: Function to check if a value is in an array of values
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Function Cleaning(text As Variant)
   Cleaning = WorksheetFunction.Clean(Trim(LCase(Unmapped_Display_Checking)))
End Function

Function Code_System_Check(unmapped_code_sys As Variant, Prev_Code_Sys As Variant) As Boolean
    Code_System_Check = InStr(Unmapped_Code_System_To_Check, Prev_Mapped_Code_Systems)
End Function


Function SortSubStrings(DelimitedString As String, Optional Delimiter As String = " ", Optional SortDescending = False) As String
    Dim SubStrings As Variant
    Dim strLeft As String, strPivot As String, strRight As String
    Dim i As Long

    SubStrings = Split(DelimitedString, Delimiter)

    If UBound(SubStrings) < 1 Then
        SortSubStrings = DelimitedString
    Else
        strPivot = SubStrings(0)

        For i = 1 To UBound(SubStrings)
            If (SubStrings(i) < strPivot) Xor SortDescending Then
                strLeft = strLeft & Delimiter & SubStrings(i)
            Else
                strRight = strRight & Delimiter & SubStrings(i)
            End If
        Next i
        strLeft = Mid(strLeft, Len(Delimiter) + 1)
        strRight = Mid(strRight, Len(Delimiter) + 1)

        If strLeft <> vbNullString Then strPivot = SortSubStrings(strLeft, Delimiter, SortDescending) & Delimiter & strPivot
        If strRight <> vbNullString Then strPivot = strPivot & Delimiter & SortSubStrings(strRight, Delimiter, SortDescending)

        SortSubStrings = strPivot
    End If
End Function
