Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
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
