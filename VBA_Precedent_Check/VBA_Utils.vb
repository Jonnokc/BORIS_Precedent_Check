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
  Dim Display As Variant
  On Error GoTo IsInArrayError:
  For Each Display In arr
    If LCase(Display) = LCase(valToBeFound) Then
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
  Next Display
IsInArrayError:
  On Error GoTo 0
  Medication_Checker = False
End Function

  ' Cleans string to remove extra spaces and put in lower case
Function Cleaning(text As Variant)
  Cleaning = WorksheetFunction.Clean(Trim(LCase(text)))
End Function

' Checks code system to see if code system is within best matches
Function Valid_Code_Sys_Checker(unmapped_code_sys As Variant, All_Valid_Code_Systems As Variant) As Boolean
  If InStr(unmapped_code_sys, ":") > 0 Then
    To_Check = Mid(unmapped_code_sys, InStr(unmapped_code_sys, ":") + 1, 200)
        If InStr(To_Check, ":") > 0 Then
            To_Check = Left(To_Check, InStr(To_Check, ":") - 1)
        End If
  Else
    To_Check = unmapped_code_sys
  End If
  For i = 1 To UBound(All_Valid_Code_Systems)
    If LCase(To_Check) = LCase(All_Valid_Code_Systems(i, 1)) Then
      Valid_Code_Sys_Checker = True
      Exit Function
    End If
  Next i
  Valid_Code_Sys_Checker = False
End Function

' Checks code system to see if code system is within best matches
Function Code_System_Check(unmapped_code_sys As Variant, Prev_Code_Sys As Variant) As Boolean
  If InStr(unmapped_code_sys, ":") > 0 Then
    To_Check = Mid(unmapped_code_sys, InStr(unmapped_code_sys, ":") + 1, 200)
  Else
    To_Check = unmapped_code_sys
  End If
  Code_System_Check = InStr(LCase(Prev_Code_Sys), LCase(To_Check))
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
On Error GoTo Text_Error
  Dim objRegex As Object
  Set objRegex = CreateObject("vbscript.regexp")
  With objRegex
    .Pattern = "()[^\w\s@]+"
    .Global = True
    ReplaceSplChars = Application.Trim(.Replace(strIn, vbNullString))
    Exit Function
  End With
  ' error handling to catch invalid displays that are just special characters
Text_Error:
    ReplaceSplChars = strIn
End Function

' Checks each word in a display against the keywords table to see if there is a match
Function Keyword_Checker(text As Variant, arr As Variant) As Variant
  SubStrings = Split(text, " ")

  For i = 1 To UBound(SubStrings)
    For j = 1 To UBound(arr)
      If SubStrings(i) = arr(j, 1) Then
        Keyword_Checker = True
        Exit Function
      End If
    Next j
  Next i
  Keyword_Checker = False
End Function

' Returns analysis of match
Function Analysis(Similarity As Variant, Keyword_Check As Variant, Code_System_Check_Answer As Variant, Medication_Check As Variant)
  If Similarity = 1 And Medication_Check = "Medication" Then
    Analysis = "Medication"
  ElseIf Similarity = 1 And Medication_Check = "Wrong Code System" Then
    Analysis = "Medication. wrong code System. Exclude"
  ElseIf Similarity = 1 And Code_System_Check_Answer = True Then
    Analysis = "Perfect Match"
  ElseIf Similarity >= 0.9 And Keyword_Check = True Then
    Analysis = "Strong Match, Includes Keyword"
  ElseIf Similarity >= 0.9 And Keyword_Check = False Then
    Analysis = "Strong Possible Match, No Keyword"
  ElseIf Similarity >= 0.85 And Keyword_Check = True Then
    Analysis = "Possible Match, Includes Keyword"
  ElseIf Keyword_Check = True Then
    Analysis = "Weak Match, Includes Keyword"
  Else
    Analysis = "Very Weak Match"
  End If
End Function

Function GetOpenFile()
  Dim fileStr As String

  MsgBox ("Could not automatically detect the Access Database.... Click OK and then select the database from the next popup.")

  File_Path = Application.GetOpenFilename()

  If GetOpenFile = "False" Then
    GetOpenFile = False
    Exit Function
  Else
    GetOpenFile = File_Path
  End If

End Function

' Tells user the status
Function Progress(Display)
  Status_Box.Show False
  With Status_Box
    .Display_Text = Display
    ' Pauses for 3 seconds so user can see the popup
    Application.Wait (Now + TimeValue("0:00:03"))
  End With

  DoEvents
End Function

' Closes the status box
Function Progress_Close()
  Unload Status_Box
End Function

Sub Del_Named_Ranges()
  For Each N In ActiveWorkbook.Names
      N.Delete
  Next
End Sub

' Disables Settings
Function Disable_Settings()
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
End Function

' Enables Settings
Function Enable_Settings()
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
End Function

' All done
Function User_Exit()
  MsgBox ("All Done! Make sure you save the file.")
  End
End Function

' Checks to see if process has run before and see if user wants to continue where left off or begin again
Function Continue_Check(Column)
  On Error GoTo All_Blank
  NextFree = Range(Column & "2:" & Column & Rows.Count).Cells.SpecialCells(xlCellTypeBlanks).Row

  If NextFree <> 2 Then
    Continue_Check_Response = MsgBox("Looks like you have ran this before. Do you want to continue from row " & NextFree & "?" & vbNewLine & vbNewLine & "Click 'Yes' to continue or 'No' to begin again.", vbYesNo + vbQuestion, "BORIS!")

    If Continue_Check_Response = vbYes Then
      Continue_Check = (NextFree - 1)
      Exit Function
    Else
      Continue_Check = 1
    End If
  Else
  All_Blank:
    Continue_Check = 1
  End If
End Function

' Checks if string is alphabetic
Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function


' Writes Data to File
Sub Write_Data(Final_All_Precedent_Closest_Match_Results, Final_Precedent_Code_System_IDs, Final_Precedent_Concept_Aliases, Final_All_Precedent_Map_Count, Final_All_Match_Responses, Final_All_Similarities)

  Range("All_Precedent_Closest_Match_Results").Value = Final_All_Precedent_Closest_Match_Results
  Range("Precedent_Code_System_Matches").Value = Final_Precedent_Code_System_IDs
  Range("Precedent_Concept_Alias_Matches").Value = Final_Precedent_Concept_Aliases
  Range("All_Precedent_Map_Count").Value = Final_All_Precedent_Map_Count
  Range("Match_Response").Value = Final_All_Match_Responses
  Range("All_Similarities").Value = Final_All_Similarities

End Sub
