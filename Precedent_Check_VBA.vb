Sub Precedence_Check()


Dim To_Review_Headers As Variant
Dim Sheet_Names As Variant
Dim Unmapped_Display_Checking As Variant
Dim Unmapped_Code_System_To_Check As Variant
Dim Prev_Mapped_Checking As Variant
Dim Prev_Mapped_Map_Count As Variant
Dim Prev_Mapped_Concept_Alias As Variant
Dim Prev_Mapped_Data_Models As Variant
Dim Prev_Mapped_Code_Systems As Variant
Dim Prev_Best_Match_Display As Variant
Dim Final_All_Unmapped_Displays_To_Check As Variant
Dim Final_All_Code_Systems_To_Check As Variant
Dim All_Prev_Mapped_Displays As Variant
Dim All_Prev_Mapped_Map_Count As Variant
Dim All_Prev_Mapped_Data_Models As Variant
Dim All_Prev_Mapped_Code_Systems As Variant
Dim All_Prev_Mapped_Concept_Aliases As Variant
Dim Header_User_Response As Variant
Dim Final_All_Precedent_Closest_Match_Results As Variant
Dim Final_All_Precedent_Closest_Match_Results_Count As Variant
Dim Final_Precedent_Code_System_IDs As Variant
Dim Final_Precedent_Concept_Aliases As Variant
Dim Final_All_Match_Responses As Variant
Dim Final_All_Similarities As Variant
Dim Final_All_Precedent_Map_Count As Variant
Dim Scrubbed_Previously_Mapped_Displays As Variant
Dim Medication_Headers As Variant

Application.ScreenUpdating = False



Sheet_Names = Array("To_Review", "Previously_Mapped", "Medications")
To_Review_Headers = Array("RAW_CODE_DISPLAY", "RAW_CODE_SYSTEM_NAME", "RAW_CODE_SYSTEM_ID", "DATA_MODEL", "Precdent Closest Match", "Precedent Raw Code System ID(s)", "Precdent Concept Alias", "MATCH_RESPONSE", "Precident Map Count", "Similarity")
Previously_Mapped_Headers = Array("Raw Code Display", "Map Count", "Precedent Data Model(s)", "Precedent Raw Code System ID(s)", "Concept Alias")
New_Headers = Array("Closest Match", "Precedent Data Model(s)", "Precedent Raw Code System ID(s)", "Precdent Concept Alias", "Similarity")
Medication_Headers = Array("drug_name")


' Checks the 'To_Review sheet to confirm the headers are all there.
  Sheets(Sheet_Names(0)).Select
  Range("A1").Select
  Range("A1", Selection.End(xlToRight)).Name = "Header_row"

    For i = 0 To UBound(To_Review_Headers)
      Header_Check = False
      For Each Header In Range("Header_row")
        If LCase(To_Review_Headers(i)) = LCase(Header) Then
          To_Review_Headers(i) = Mid(Header.Address, 2, 1)
          Header_Check = True
          Exit For
        End If
      Next Header
      If Header_Check = False Then
        Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & To_Review_Headers(i) & "'" & " on the " & Sheet_Names(0) & " sheet. Enter the letter of that column in the text box below to continue the program. Or click cancel to exit program.")

        'If user hits cancel then close program.
        If Header_User_Response = vbNullString Then
          GoTo User_Exit
        Else
          ' enters the title of the column the user just identified
          Range(Header_User_Response & "1").Value = To_Review_Headers(i)
          To_Review_Headers(i) = Header_User_Response
        End If
      End If
    Next i


    ' Checks and stores all the header locations for the Previously mapped sheet.
    Sheets(Sheet_Names(1)).Select
    Range("A1").Select
    Range("A1", Selection.End(xlToRight)).Name = "Header_row"

      For i = 0 To UBound(Previously_Mapped_Headers)
        Header_Check = False
        For Each Header In Range("Header_row")
          If LCase(Previously_Mapped_Headers(i)) = LCase(Header) Then
            Previously_Mapped_Headers(i) = Mid(Header.Address, 2, 1)
            Header_Check = True
            Exit For
          End If
        Next Header
        If Header_Check = False Then
          Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Previously_Mapped_Headers(i) & "'" & " on the " & Sheet_Names(1) & " sheet. Enter the letter of that column in the text box below to continue the program. Or click cancel to exit program.")

          'If user hits cancel then close program.
          If Header_User_Response = vbNullString Then
            GoTo User_Exit
          Else
            ' enters the title of the column the user just identified
            Range(Header_User_Response & "1").Value = Previously_Mapped_Headers(i)
            Previously_Mapped_Headers(i) = Header_User_Response
          End If
        End If
      Next i



' SUB - Saves all the data to Memory
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Names the currently unmapped proprietary code displays as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(0) & "2:" & To_Review_Headers(0) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "All_Unmapped_Displays_To_Check"
'Saves range to array
Final_All_Unmapped_Displays_To_Check = Range("All_Unmapped_Displays_To_Check").Value

' Names the raw code systems to check column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(1) & "2:" & To_Review_Headers(1) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "All_Raw_Code_System_Names_To_Check"
'Saves range to array
Final_All_Code_Systems_To_Check = Range("All_Raw_Code_System_Names_To_Check").Value

' Names the final results closest match column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(4) & "2:" & To_Review_Headers(4) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "All_Precedent_Closest_Match_Results"
'Saves range to array
Final_All_Precedent_Closest_Match_Results = Range("All_Precedent_Closest_Match_Results").Value

' Names the final results map count column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(8) & "2:" & To_Review_Headers(8) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "All_Precedent_Map_Count"
'Saves range to array
Final_All_Precedent_Map_Count = Range("All_Precedent_Map_Count").Value

' Names the final results code system id column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(5) & "2:" & To_Review_Headers(5) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Precedent_Code_System_Matches"
'Saves range to array
Final_Precedent_Code_System_IDs = Range("Precedent_Code_System_Matches").Value

' Names the final results concept alias column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(6) & "2:" & To_Review_Headers(6) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Precedent_Concept_Alias_Matches"
'Saves range to array
Final_Precedent_Concept_Aliases = Range("Precedent_Concept_Alias_Matches").Value

' Names the Match Response column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(7) & "2:" & To_Review_Headers(7) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Match_Response"
'Saves range to array
Final_All_Match_Responses = Range("Match_Response").Value

' Names the similarity column as a range
Sheets(Sheet_Names(0)).Range(To_Review_Headers(9) & "2:" & To_Review_Headers(9) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "All_Similarities"
'Saves range to array
Final_All_Similarities = Range("All_Similarities").Value

' Names the already validated raw code displays as a range
Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(0) & "2:" & Previously_Mapped_Headers(0) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Previously_Mapped_Displays"
'Saves range to array
All_Prev_Mapped_Displays = Range("Previously_Mapped_Displays").Value

' Names the Previously mapped map count as a range
Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(1) & "2:" & Previously_Mapped_Headers(1) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Previously_Mapped_Map_Count"
'Saves range to array
All_Prev_Mapped_Map_Count = Range("Previously_Mapped_Map_Count").Value

' Names the previously mapped data models as a range
Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(2) & "2:" & Previously_Mapped_Headers(2) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Previously_Mapped_Data_Models"
'Saves range to array
All_Prev_Mapped_Data_Models = Range("Previously_Mapped_Data_Models").Value

' Names the previously mapped code systems ids as a range
Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(3) & "2:" & Previously_Mapped_Headers(3) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Previously_Mapped_Code_Systems"
'Saves range to array
All_Prev_Mapped_Code_Systems = Range("Previously_Mapped_Code_Systems").Value

' Names the previously mapped concept aliases as a range
Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(4) & "2:" & Previously_Mapped_Headers(4) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "Previously_Mapped_Concept_Alias"
'Saves range to array
All_Prev_Mapped_Concept_Aliases = Range("Previously_Mapped_Concept_Alias").Value


' Names the Medication displays as a Range
Sheets(Sheet_Names(2)).Range("A" & "2:" & "A" & Sheets(Sheet_Names(2)).Cells.SpecialCells(xlCellTypeLastCell).row).Name = "All_Medication_Displays"
' Saves range to array
All_Prev_Medication_Displays = Range("All_Medication_Displays")



Sheets(Sheet_Names(0)).Select

' SUB - SCRUBS PREVIOUSLY MAPPED DISPLAYS TO STANDARDIZE FORMAT
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 1 To UBound(All_Prev_Mapped_Displays)
  All_Prev_Mapped_Displays(i, 1) = Utils.Cleaning(All_Prev_Mapped_Displays(i, 1))
Next i

' TODO - Loop through Previously mapped and currently mapped and sort words in string alphabetically


' SUB - ACTUAL LOGIC TO FIND MATCHES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

For i = 1 To UBound(Final_All_Unmapped_Displays_To_Check)

  ' Cell value currently looking for a closest match with
  Unmapped_Display_Checking = Final_All_Unmapped_Displays_To_Check(i, 1)
  Unmapped_Code_System_To_Check = Final_All_Code_Systems_To_Check(i, 1)
  Best_Similarity = 0
  Code_System_Check_Answer = False

  DoEvents
  Application.StatusBar = "Current progress: " & (i / UBound(Final_All_Unmapped_Displays_To_Check) * 100) & "% complete"


    ' Checks the currently unmapped against all previously unmapped
  For j = 1 To UBound(All_Prev_Mapped_Displays)
    Prev_Mapped_Checking = All_Prev_Mapped_Displays(j, 1)

    ' Medication Checks
    Medication_Check = Utils.IsInArray(Prev_Mapped_Checking, All_Prev_Medication_Displays)

    Fuzzy_Similarity = Functions.HotFuzz(Unmapped_Display_Checking, Prev_Mapped_Checking)
    ' Checks the code system and awards a bonus for a match


    If Fuzzy_Similarity = 1 Then
      best_match_index = j
      Best_Similarity = Fuzzy_Similaritya
      Exit For

    ElseIf Best_Similarity > Fuzzy_Similarity Then
      ' Saves the values of the current best match
      best_match_index = j
      Best_Similarity = Fuzzy_Similarity
    End If
  Next j

  If Best_Similarity > 0.5 Then
    ' Checks if code system is within best match
    Code_System_Check_Answer = Code_System_Check(Unmapped_Code_System_To_Check, All_Prev_Mapped_Code_Systems(best_match_index, 1))

    Final_All_Precedent_Closest_Match_Results(i, 1) = All_Prev_Mapped_Displays(best_match_index, 1)
    Final_Precedent_Code_System_IDs(i, 1) = All_Prev_Mapped_Code_Systems(best_match_index, 1)
    Final_Precedent_Concept_Aliases(i, 1) = All_Prev_Mapped_Concept_Aliases(best_match_index, 1)
    Final_All_Precedent_Map_Count(i, 1) = All_Prev_Mapped_Map_Count(best_match_index, 1)
    Final_All_Similarities(i, 1) = Best_Similarity
    If Code_System_Check_Answer = False Then
      Final_All_Match_Responses(i, 1) = "Code System of Closest Match IS NOT Within Previously Validated"
    ElseIf Code_System_Check_Answer = True Then
      Final_All_Match_Responses(i, 1) = "Code System of Closest Match IS Within Previously Validated"
    End If
  End If
Next i

User_Exit:

' Writes the results back to the table.
Range("All_Precedent_Closest_Match_Results").Value = Final_All_Precedent_Closest_Match_Results
Range("Precedent_Code_System_Matches").Value = Final_Precedent_Code_System_IDs
Range("Precedent_Concept_Alias_Matches").Value = Final_Precedent_Concept_Aliases
Range("All_Precedent_Map_Count").Value = Final_All_Precedent_Map_Count
Range("Match_Response").Value = Final_All_Match_Responses
Range("All_Similarities").Value = Final_All_Similarities


Application.StatusBar = False
Application.ScreenUpdating = True

End Sub
