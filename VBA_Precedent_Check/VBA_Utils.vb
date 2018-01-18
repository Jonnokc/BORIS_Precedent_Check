Option Private Module

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
  Dim i       As Long

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
  Dim intPos  As Integer
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


' Deletes extra sheets / Cleanup
Sub Delete_Extra_Sheets()

  Dim sheet   As Worksheet

  Application.DisplayAlerts = False

  For Each sheet In Worksheets

    If sheet.Name = "Medications" _
       Or sheet.Name = "Keywords" _
       Or sheet.Name = "Valid_Code_Systems" _
       Or sheet.Name = "Previously_Mapped" _
       Then
      sheet.Delete
    End If
  Next sheet

  Application.DisplayAlerts = True

End Sub


' Adds color rule
Sub color(Similarity_Column)

  Columns(Similarity_Column & ":" & Similarity_Column).Select
  Selection.FormatConditions.AddColorScale ColorScaleType:=3
  Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
  Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
                                                           xlConditionValueNumber
  Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 0
  With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
    .color = 7039480
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
                                                           xlConditionValueNumber
  Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0.5
  With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
    .color = 8711167
    .TintAndShade = 0
  End With
  Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
                                                           xlConditionValueNumber
  Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 1
  With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
    .color = 8109667
    .TintAndShade = 0
  End With

End Sub


' Deletes conditional formatting
Sub Delete_CF(Sht, Column)

  With Sht.Range(Column & ":" & Column)
    .FormatConditions.Delete
  End With

End Sub


' Converts All Non Formula Cells To General Format
Function Convert_Non_Formula_To_General(Worksheet)
  Dim rng     As Range
  Dim cl      As Range
  Dim rConst  As Range

  Sheets(Worksheet).Select

  ' pick an unused cell
  Set rConst = Cells(27, 27)
  rConst = 1

  ' Formats all non formula cells as "General"
  Set rng = Cells.SpecialCells(xlCellTypeConstants)
  rng.NumberFormat = "General"
  rConst.Copy
  rng.PasteSpecial xlPasteValues, xlPasteSpecialOperationMultiply

  rConst.Clear

End Function


' Converts All Non Formula Cells To General Format
Function Convert_All_Visible_To_General(Worksheet)
  Dim rng     As Range
  Dim Sht     As Worksheet

  Sheets(Worksheet).Select
  Set Sht = Worksheets(Worksheet)
  Set StartCell = Range("A1")

  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column



  ' Formats all non formula cells as "General"
  Set rng = Range(StartCell, Sht.Cells(LastRow, LastColumn))
  rng.Copy
  Selection.PasteSpecial paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                         :=False, Transpose:=False

  ' Clears the copy selection outline
  Application.CutCopyMode = False

  ' Keep things clean
  Range("A1").Select

End Function


' Formats range as a table
Function Format_As_Table(Table_Name, Worksheet)
  Dim Sht     As Worksheet

  ' Formats the data on the "To Review" Sheet into a TableStyle
  Sheets(Worksheet).Select
  Set Sht = Worksheets(Worksheet)
  Set StartCell = Range("A1")

  If Sht.AutoFilterMode = True Then
    Sht.AutoFilterMode = False
  End If

  ' Checks the current sheet. If it is in table format, convert it to range.
  If Sht.ListObjects.Count > 0 Then
    With Sht.ListObjects(1)
      Set rList = .Range
      .Unlist
    End With
  End If


  'Find Last Row and Column
  LastRow = StartCell.SpecialCells(xlCellTypeLastCell).Row
  LastColumn = StartCell.SpecialCells(xlCellTypeLastCell).Column

  'Select Range
  Sht.Range(StartCell, Sht.Cells(LastRow, LastColumn)).Select

  ' Clears all formats on the cells before making them a table
  Selection.ClearFormats

  'Turn selected Range Into Table
  Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
  tbl.Name = Table_Name
  tbl.TableStyle = "TableStyleLight12"
  Cells.Select
  ' To keep things tidy
  Range("A1").Select

End Function


' Finds and stores locations of Columns
Function Find_Col_Location(Worksheet, Header_Array, Header_Num_Array, Header_Row_Start_Cell)
  Sheets(Worksheet).Select
  Range(Header_Row_Start_Cell).Select
  Range(Header_Row_Start_Cell, Selection.End(xlToRight)).Name = "Header_row"

  For i = 0 To UBound(Header_Array)
    Header_Check = False
    For Each Header In Range("Header_row")
      If LCase(Header_Array(i)) = LCase(Header) Then
        Header_Array(i) = Mid(Header.Address, 2, 1)
        Header_Num_Array(i) = Range(Header_Array(i) & "1").Column
        Header_Check = True
        Exit For
      End If
    Next Header
    If Header_Check = False Then
      Call Utils.Enable_Settings
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Header_Array(i) & "'" & " on the " & Worksheet & " sheet. Enter the letter of that column in the text box below to continue the program. Or click cancel to exit program.")

      ' If user hits cancel then close program.
      If Header_User_Response = vbNullString Then
        Call Error_Handling.Exiting_User_Error("No column letter provided. Exiting Program")
      Else
        ' enters the title of the column the user just identified
        Range(Header_User_Response & "1").Value = Header_Array(i)
        Header_Array(i) = Header_User_Response
        Header_Num_Array(i) = Range(Header_Array(i) & "1").Column
        Call Utils.Disable_Settings
      End If
    End If
  Next i
End Function


' Deletes visible rows from Table Single Criteria
Function Delete_Visible_Rows_Single_Criteria(Worksheet, List_Object_Num, Filter_Col_1, Filter_1_Criteria_1)
  Dim RngToDelete As Range
  Dim TabelName As ListObject
  Dim L       As ListObject



  ' filters to just show validated
  Sheets(Worksheet).Select

  ' filters for just "all validated rows to delete them"
  Sheets(Worksheet).ListObjects(List_Object_Num).Range.AutoFilter Field:=Filter_Col_1, Criteria1 _
                                                                  :=Filter_1_Criteria_1


  ' Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
  Set tbl = Sheets(Worksheet).ListObjects(List_Object_Num)

  If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
    Table_ObjIsVisible = True
  Else:
  Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
End If

' If deletes all rows that are "all validated"
If Table_ObjIsVisible = True Then
  Sheets(Worksheet).Select

  Set TabelName = ActiveSheet.ListObjects(List_Object_Num)
  With TabelName.Range
    Set RngToDelete = .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    .AutoFilter
    RngToDelete.Delete
  End With
End If

' Clears all filters if they are on.
For Each L In ActiveSheet.ListObjects
  'Show all data if the AutoFilter of the table is enabled
  If L.ShowAutoFilter Then
    If L.AutoFilter.FilterMode Then
      L.AutoFilter.ShowAllData
    End If
  End If
Next L


End Function


' Deletes visible rows from Table Two Criteria
Function Delete_Visible_Rows_Two_Criteria(Worksheet, List_Object_Num, Filter_Col_1, Filter_1_Criteria_1, Filter_Col_2, Filter_Col_2_Criteria_1, Filter_Col_2_Criteria_2)
  Dim RngToDelete As Range
  Dim TabelName As ListObject


  ' filters to remove validated for PowerForms, and iView and validated codes
  Sheets(Worksheet).ListObjects(List_Object_Num).Range.AutoFilter Field:=Filter_Col_1, Criteria1 _
                                                                  :=Filter_1_Criteria_1

  Sheets(Worksheet).ListObjects(List_Object_Num).Range.AutoFilter Field:=Filter_Col_2, _
                                                                  Criteria1:=Filter_Col_2_Criteria_1, Operator:=xlOr, _
                                                                  Criteria2:=Filter_Col_2_Criteria_2


  ' Checks current table to determine if any cells are visible. If cells are visible then set "Table_ObjIsVisible" = TRUE
  Set tbl = Sheets(Worksheet).ListObjects(List_Object_Num)

  If tbl.Range.SpecialCells(xlCellTypeVisible).Areas.Count > 1 Then
    Table_ObjIsVisible = True
  Else:
  Table_ObjIsVisible = tbl.Range.SpecialCells(xlCellTypeVisible).Rows.Count > 1
End If

' If deletes all rows that are "all validated"
If Table_ObjIsVisible = True Then
  Sheets(Worksheet).Select

  Set TabelName = ActiveSheet.ListObjects(List_Object_Num)
  With TabelName.Range
    Set RngToDelete = .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    .AutoFilter
    RngToDelete.Delete
  End With
End If

' Clears all filters if they are on.
' Clears all filters if they are on.
For Each L In ActiveSheet.ListObjects
  'Show all data if the AutoFilter of the table is enabled
  If L.ShowAutoFilter Then
    If L.AutoFilter.FilterMode Then
      L.AutoFilter.ShowAllData
    End If
  End If
Next L

End Function
