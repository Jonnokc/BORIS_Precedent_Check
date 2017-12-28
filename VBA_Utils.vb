'DESCRIPTION: Function to check if a value is in an array of values
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
  Dim element As Variant
  On Error GoTo IsInArrayError:                  'array is empty
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

' Checks if unmapped code is on the medication list.
Function Medication_Checker(valToBeFound As Variant, arr As Variant, code_system As Variant)
  Dim display As Variant
  On Error GoTo IsInArrayError:
  For Each display In arr
    If LCase(display) Like "*" & LCase(valToBeFound) & "*" Then
      If LCase(code_system) Like "*medication*" Or _
      LCase(code_system) Like "*drug*" Or _
      LCase(code_system) Like "*allergy*" Then
        Medication_Checker = "Medication"
        Exit Function
      Else
        Medication_Checker = "Wrong Code System"
        Exit Function
      End If
    End If
  Next display
IsInArrayError:
  On Error GoTo 0
  Medication_Checker = False
End Function

  ' Cleans string to remove extra spaces and put in lower case
Function Cleaning(text As Variant)
  Cleaning = WorksheetFunction.Clean(Trim(LCase(text)))
End Function

  ' Checks code system to see if code system is within best matches
Function Code_System_Check(unmapped_code_sys As Variant, Prev_Code_Sys As Variant) As Boolean
  Code_System_Check = InStr(Unmapped_Code_System_To_Check, Prev_Mapped_Code_Systems)
End Function

' Sorts all words in string alphabetically
Function Sort_Sub_Strings(DelimitedString As Variant, Optional Delimiter As String = " ", Optional SortDescending = False) As String
  Dim SubStrings As Variant
  Dim strLeft As String, strPivot As String, strRight As String
  Dim i As Long

  SubStrings = Split(DelimitedString, Delimiter)

  If UBound(SubStrings) < 1 Then
    Sort_Sub_Strings = DelimitedString
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

    If strLeft <> vbNullString Then strPivot = Sort_Sub_Strings(strLeft, Delimiter, SortDescending) & Delimiter & strPivot
    If strRight <> vbNullString Then strPivot = strPivot & Delimiter & Sort_Sub_Strings(strRight, Delimiter, SortDescending)

    Sort_Sub_Strings = strPivot
  End If
End Function

' Replaces the special characters with nothing.
Function ReplaceSplChars(strIn As Variant) As String
  Dim objRegex As Object
  Set objRegex = CreateObject("vbscript.regexp")
  With objRegex
    .Pattern = "()[^\w\s@]+"
    .Global = True
    ReplaceSplChars = Application.Trim(.Replace(strIn, vbNullString))
  End With
End Function

' Checks each word in a display against the keywords table to see if there is a match
Function Keyword_Checker(text As Variant, arr As Variant) As Variant
  SubStrings = Split(text, " ")

  For i = 1 To UBound(SubStrings)
    For j = 1 To UBound(arr)
      If SubStrings(i) = arr(j, 1) Then
        Keyword_Checker = True
      End If
    Next j
  Next i
  Keyword_Checker = False
End Function

Function Analysis(Similarity As Variant, Keyword_Check As Variant, Code_System_Check_Answer As Variant, Medication_Check As Variant)
  If Similarity = 1 And Code_System_Check_Answer = True Then
    Analysis = "Perfect Match"
  ElseIf Similarity >= 0.9 And Keyword_Check = True Then
    Analysis = "Strong Match, Keyword"
  ElseIf Similarity = 1 And Medication_Check = "Medication" Then
    Analysis = "Medication"
  ElseIf Similarity = 1 And Medication_Check = "Wrong Code System" then
    Analysis = "Medication display with wrong code System. Exclude"
  ElseIf Similarity >= 0.7 And Keyword_Check = True Then
    Analysis = "Possible Match, Keyword"
  Else
    Analysis = "Weak Match"
  End If
End Function
