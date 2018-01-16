Public Abort As Boolean

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
  Dim Keywords_Headers As Variant
  Dim Valid_Code_Sys_Headers As Variant
  Dim LastRow As Long
  Dim pcstdone As Single
  Dim Start_Row As Long


  ' Disables settings to improve macro performance.
  Call Utils.Disable_Settings


  Sheet_Names = Array("To_Review", "Previously_Mapped", "Medications", "Keywords", "Valid_Code_Systems")
  To_Review_Headers = Array("RAW_CODE_DISPLAY", "RAW_CODE_SYSTEM_NAME", "RAW_CODE_SYSTEM_ID", "Precedent Closest Match", "Precedent Raw Code System ID(s)", "Precedent Concept Alias", "MATCH_RESPONSE", "Precedent Map Count", "Similarity")
  Previously_Mapped_Headers = Array("Raw Code Display", "Map Count", "Precedent Raw Code System ID(s)", "Concept Alias")
  New_Headers = Array("Closest Match", "Precedent Raw Code System ID(s)", "Precedent Concept Alias", "Similarity")
  Medication_Headers = Array("drug_name")
  Keywords_Headers = Array("Keywords")
  Valid_Code_Sys_Headers = Array("Code_Systems")


  ' SUB - IMPORTS DATA FROM DATABASE
  '''''''''''''''''''''''''''''''''''''''''

  ' Imports the data from the database
  Call Get_data.Data_Import

  ' Deletes named ranges if they exist
  Call Utils.Del_Named_Ranges



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
      Call Utils.Enable_Settings

      ' If the header not found is one of the new precedent check headers, then just add it and save location
      If i >= 3 Then
        Add_Header = Mid(Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1).Address, 2, 1)
        Range(Add_Header & "1") = To_Review_Headers(i)
        To_Review_Headers(i) = Add_Header
      Else
        Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & To_Review_Headers(i) & "'" & " on the " & Sheet_Names(1) & " sheet. Enter the letter of that column in the text box below to continue the program. Or click cancel to exit program.")

        ' If user hits cancel then close program.
        If Header_User_Response = vbNullString Then
          GoTo User_Exit
        Else
          ' enters the title of the column the user just identified
          Range(Header_User_Response & "1").Value = To_Review_Headers(i)
          To_Review_Headers(i) = Header_User_Response
          Call Utils.Disable_Settings
        End If
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
      Call Utils.Enable_Settings
      Header_User_Response = InputBox("BORIS was unable to find the header:" & vbNewLine & "'" & Previously_Mapped_Headers(i) & "'" & " on the " & Sheet_Names(1) & " sheet. Enter the letter of that column in the text box below to continue the program. Or click cancel to exit program.")

      ' If user hits cancel then close program.
      If Header_User_Response = vbNullString Then
        GoTo User_Exit
      Else
        ' enters the title of the column the user just identified
        Range(Header_User_Response & "1").Value = Previously_Mapped_Headers(i)
        Previously_Mapped_Headers(i) = Header_User_Response
        Call Utils.Disable_Settings
      End If
    End If
  Next i


' SUB - Saves all the data to Memory
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' Data range error handling
  On Error GoTo Range_Error

  ' Names the currently unmapped proprietary code displays as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(0) & "2:" & To_Review_Headers(0) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Unmapped_Displays_To_Check"
  ' Saves range to array
  Final_All_Unmapped_Displays_To_Check = Range("All_Unmapped_Displays_To_Check").Value

  ' Names the raw code systems to check column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(1) & "2:" & To_Review_Headers(1) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Raw_Code_System_Names_To_Check"
  ' Saves range to array
  Final_All_Code_Systems_To_Check = Range("All_Raw_Code_System_Names_To_Check").Value

  ' Names the final results closest match column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(3) & "2:" & To_Review_Headers(3) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Precedent_Closest_Match_Results"
  ' Saves range to array
  Final_All_Precedent_Closest_Match_Results = Range("All_Precedent_Closest_Match_Results").Value

  ' Names the final results map count column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(7) & "2:" & To_Review_Headers(7) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Precedent_Map_Count"
  ' Saves range to array
  Final_All_Precedent_Map_Count = Range("All_Precedent_Map_Count").Value

  ' Names the final results code system id column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(4) & "2:" & To_Review_Headers(4) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Precedent_Code_System_Matches"
  ' Saves range to array
  Final_Precedent_Code_System_IDs = Range("Precedent_Code_System_Matches").Value

  ' Names the final results concept alias column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(5) & "2:" & To_Review_Headers(5) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Precedent_Concept_Alias_Matches"
  ' Saves range to array
  Final_Precedent_Concept_Aliases = Range("Precedent_Concept_Alias_Matches").Value

  ' Names the Match Response column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(6) & "2:" & To_Review_Headers(6) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Match_Response"
  ' Saves range to array
  Final_All_Match_Responses = Range("Match_Response").Value

  ' Names the similarity column as a range
  Sheets(Sheet_Names(0)).Range(To_Review_Headers(8) & "2:" & To_Review_Headers(8) & Sheets(Sheet_Names(0)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Similarities"
  ' Saves range to array
  Final_All_Similarities = Range("All_Similarities").Value

  ' Names the already validated raw code displays as a range
  Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(0) & "2:" & Previously_Mapped_Headers(0) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Previously_Mapped_Displays"
  ' Saves range to array
  All_Prev_Mapped_Displays = Range("Previously_Mapped_Displays").Value
  ' used for storing scrubbed version
  Scrubbed_All_Prev_Mapped_Displays = Range("Previously_Mapped_Displays").Value

  ' Names the Previously mapped map count as a range
  Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(1) & "2:" & Previously_Mapped_Headers(1) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Previously_Mapped_Map_Count"
  ' Saves range to array
  All_Prev_Mapped_Map_Count = Range("Previously_Mapped_Map_Count").Value

  ' Names the previously mapped code systems ids as a range
  Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(2) & "2:" & Previously_Mapped_Headers(2) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Previously_Mapped_Code_Systems"
  ' Saves range to array
  All_Prev_Mapped_Code_Systems = Range("Previously_Mapped_Code_Systems").Value

  ' Names the previously mapped concept aliases as a range
  Sheets(Sheet_Names(1)).Range(Previously_Mapped_Headers(3) & "2:" & Previously_Mapped_Headers(3) & Sheets(Sheet_Names(1)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "Previously_Mapped_Concept_Alias"
  ' Saves range to array
  All_Prev_Mapped_Concept_Aliases = Range("Previously_Mapped_Concept_Alias").Value

  ' Names the Medication displays as a Range
  Sheets(Sheet_Names(2)).Range("A" & "2:" & "A" & Sheets(Sheet_Names(2)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Medication_Displays"
  ' Saves range to array
  All_Prev_Medication_Displays = Range("All_Medication_Displays")

  ' Names the Keywords displays as a Range
  Sheets(Sheet_Names(3)).Range("A" & "2:" & "A" & Sheets(Sheet_Names(3)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Keywords_Displays"
  ' Saves range to array
  All_Prev_Keywords_Displays = Range("All_Keywords_Displays")

  ' Names the valid code systems as a range
  Sheets(Sheet_Names(4)).Range("A" & "2:" & "A" & Sheets(Sheet_Names(4)).Cells.SpecialCells(xlCellTypeLastCell).Row).Name = "All_Valid_Code_Systems"
  ' Saves range to array
  All_Valid_Code_Systems = Range("All_Valid_Code_Systems")

  ' just for cleaning
  Sheets(Sheet_Names(0)).Select


' SUB - CHECKS TO SEE IF PROCESS HAS RUN BEFORE AND CHECKS WITH USER TO CONTINUE OR BEGIN AGAIN.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' resets error handling
  On Error GoTo 0

  ' Finds start row if process has been run before and ended early
  Start_Row = Utils.Continue_Check(To_Review_Headers(6))


' SUB - SCRUBS PREVIOUSLY MAPPED DISPLAYS TO STANDARDIZE FORMAT
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ufProgress.LabelProgress.Width = 0
  ' forces progress bar to screen
  ufProgress.Show
  For i = Start_Row To UBound(Scrubbed_All_Prev_Mapped_Displays)

    ' User exiting program early
    If Abort = True Then
      GoTo User_Exit
    End If

    ' Sets error handling to go to next row on an error
    On Error Resume Next

    pcstdone = i / UBound(Scrubbed_All_Prev_Mapped_Displays)
    ' Cleans format, removes special characters, orders string alphabetically
    Scrubbed_All_Prev_Mapped_Displays(i, 1) = Utils.Sort_Sub_Strings(Utils.Cleaning(Utils.ReplaceSplChars(Scrubbed_All_Prev_Mapped_Displays(i, 1))))
    ' updates progress bar
    With ufProgress
      .LabelCaption = "Cleaning Data... Processing Row " & i & " of " & UBound(Scrubbed_All_Prev_Mapped_Displays)
      .LabelProgress.Width = pcstdone * (.FrameProgress.Width)
    End With
    DoEvents
    ' closes progress bar when done
    If i = UBound(Scrubbed_All_Prev_Mapped_Displays) Then
      Unload ufProgress
    End If
  Next i


' SUB - ACTUAL LOGIC TO FIND MATCHES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Calls status update for user
  Call Utils.Progress("Getting Ready to check codes....")
  Call Utils.Progress_Close

  ' Changes project bar settings
  ufProgress.LabelProgress.Width = 0
  ufProgress.LabelProgress.BackColor = &H80FF&
  ufProgress.Show

  ' Loops through unmapped codes
  For i = 1 To UBound(Final_All_Unmapped_Displays_To_Check)
    ' Traditional Error handling. Skip row if there is an error
    On Error Resume Next

    pcstdone = i / UBound(Final_All_Unmapped_Displays_To_Check)

    With ufProgress
      .LabelCaption = "Processing Data....You see me I be work, work, work, work, work.... " & i & " of " & UBound(Final_All_Unmapped_Displays_To_Check)
      .LabelProgress.Width = pcstdone * (.FrameProgress.Width)
    End With
    DoEvents

    ' Cell value currently looking for a closest match with
    Unmapped_Display_Checking = Utils.Sort_Sub_Strings(Utils.Cleaning(Utils.ReplaceSplChars(Final_All_Unmapped_Displays_To_Check(i, 1))))
    Unmapped_Code_System_To_Check = Final_All_Code_Systems_To_Check(i, 1)
    Best_Similarity = 0

    ' Checks the code system of the current code to make sure it is a code system that PCST maps.
    Legit_Code_Sys_Check = Utils.Valid_Code_Sys_Checker(Unmapped_Code_System_To_Check, All_Valid_Code_Systems)

    ' If code system is is valid, then continue, else go to next
    If Legit_Code_Sys_Check = True Then

      ' Medication Checker. Looks to see if code is a legit medication, or not a medication at all.
      Medication_Check = Utils.Medication_Checker(Unmapped_Display_Checking, All_Prev_Medication_Displays, Unmapped_Code_System_To_Check)

      ' Keyword Checker. Checks each word of display against keyword table.
      Keyword_Check = Utils.Keyword_Checker(Unmapped_Display_Checking, All_Prev_Keywords_Displays)

      ' If display display is not on the medication list or IS with the correct code system
      If (Medication_Check = False Or Medication_Check = "Medication") And Legit_Code_Sys_Check = True Then

        ' Checks the currently unmapped against all previously unmapped
        For j = 1 To UBound(Scrubbed_All_Prev_Mapped_Displays)

          Prev_Mapped_Checking = Scrubbed_All_Prev_Mapped_Displays(j, 1)

          Fuzzy_Similarity = Functions.HotFuzz(Unmapped_Display_Checking, Prev_Mapped_Checking)

          If Fuzzy_Similarity = 1 Then
            best_match_index = j
            Best_Similarity = Fuzzy_Similarity
            Exit For

          ElseIf Fuzzy_Similarity > Best_Similarity Then
            ' Saves the values of the current best match
            best_match_index = j
            Best_Similarity = Fuzzy_Similarity
          End If

        Next j

      End If

    End If


' SUB Determines best values and saves to Array
''''''''''''''''''''''''''''''''''''''''''''''''

    If Best_Similarity >= 0.7 Or Keyword_Check = True Then
      ' Resets error handling to default
      On Error GoTo 0

      ' Checks if code system is within best match
      Code_System_Check_Answer = Code_System_Check(Unmapped_Code_System_To_Check, All_Prev_Mapped_Code_Systems(best_match_index, 1))
      ' determines the response based on analysis
      Final_All_Match_Responses(i, 1) = Analysis(Best_Similarity, Keyword_Check, Code_System_Check_Answer, Medication_Check)
      ' populates fields with array based on best matches
      Final_All_Precedent_Closest_Match_Results(i, 1) = All_Prev_Mapped_Displays(best_match_index, 1)
      Final_Precedent_Code_System_IDs(i, 1) = All_Prev_Mapped_Code_Systems(best_match_index, 1)
      Final_Precedent_Concept_Aliases(i, 1) = All_Prev_Mapped_Concept_Aliases(best_match_index, 1)
      Final_All_Precedent_Map_Count(i, 1) = All_Prev_Mapped_Map_Count(best_match_index, 1)
      Final_All_Similarities(i, 1) = Best_Similarity

    ' Not a valid code system. Nothing to write.
    ElseIf Legit_Code_Sys_Check = False Then
      Final_All_Match_Responses(i, 1) = "Bad Code System. PCST Does Not Map"
      Final_All_Precedent_Closest_Match_Results(i, 1) = ""
      Final_Precedent_Code_System_IDs(i, 1) = ""
      Final_Precedent_Concept_Aliases(i, 1) = ""
      Final_All_Precedent_Map_Count(i, 1) = ""
      Final_All_Similarities(i, 1) = ""

    ' No match and does not contain a keyword
    Else
      Final_All_Match_Responses(i, 1) = "No Match Found"
      Final_All_Precedent_Closest_Match_Results(i, 1) = ""
      Final_Precedent_Code_System_IDs(i, 1) = ""
      Final_Precedent_Concept_Aliases(i, 1) = ""
      Final_All_Precedent_Map_Count(i, 1) = ""
      Final_All_Similarities(i, 1) = ""
    End If

    If i = UBound(Final_All_Unmapped_Displays_To_Check) Then
      Unload ufProgress
    End If

  Next i

User_Exit:

' TODO - add error handling so if user cancels that current progress is written to file.

  ' Writes the results back to the table.

  Call Utils.Write_Data(Final_All_Precedent_Closest_Match_Results, Final_Precedent_Code_System_IDs, Final_Precedent_Concept_Aliases, Final_All_Precedent_Map_Count, Final_All_Match_Responses, Final_All_Similarities)

  ' Range("All_Precedent_Closest_Match_Results").Value = Final_All_Precedent_Closest_Match_Results
  ' Range("Precedent_Code_System_Matches").Value = Final_Precedent_Code_System_IDs
  ' Range("Precedent_Concept_Alias_Matches").Value = Final_Precedent_Concept_Aliases
  ' Range("All_Precedent_Map_Count").Value = Final_All_Precedent_Map_Count
  ' Range("Match_Response").Value = Final_All_Match_Responses
  ' Range("All_Similarities").Value = Final_All_Similarities

  ' Program done.
  Call Utils.Enable_Settings
  Call Utils.User_Exit

Range_Error:
  Call Utils.Enable_Settings
  Call Error_Handling.Range_Error

End Sub
