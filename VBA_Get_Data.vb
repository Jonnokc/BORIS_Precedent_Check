Sub Get_Data()

Dim sheet As Worksheet
Dim Sheet_Names As Variant
Dim Access_Tbl_Names As Variant
Dim Access_File_Path As Variant



Sheet_Names = Array("Previously_Mapped", "Medications", "Keywords", "Valid_Code_Systems")
Access_Tbl_Names = Array("Q_Raw_Codes_To_Concept_Alias_2017", "Q_All_Medication_Raw_Code_Displays_Excel", "Proprietary_Code_Display_Keywords", "Mappable_Raw_Code_System_Names")

Access_File_Path = "Y:\Data Intelligence\Code_Submittion_Database\CodeFeedbackDatabase.accdb"


' SUB - Checks to confirm user has path to database mapped correctly. If not, prompt user for file path.

' SUB - Checks to confirm path to Access Database is mapped correctly. If not then it asks the user to get it.
Set fso = CreateObject("scripting.filesystemobject")
With fso
  Path_Checker = Len(Dir("Y:\Data Intelligence\Code_Database\Data_Intelligence_Code_Database.accdb")) <> 0
End With

If Path_Checker = 0 Then
  Access_Database_Check = Utils.GetOpenFile()
  If Access_Database_Check = False Then
    GoTo User_Exit
  Else
    Access_File_Path = Access_Database_Check
  End If
Else
  ' Do Nothing
End If
Set fso = Nothing


' SUB - Delets all the extra sheets and then readds them to make sure things remain clean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ' Deletes extra sheet
  Application.DisplayAlerts = False

  For Each sheet In Worksheets
      If sheet.Name = "Medications" _
      Or sheet.Name = "Keywords" _
      Or sheet.Name = "Previously_Mapped" _
      Or sheet.Name = "Valid_Code_Systems" _
      Then
        sheet.Delete
      End If
  Next sheet

  Application.DisplayAlerts = True

  ' Creates the sheets within the workbook
  For i In UBound(Sheet_Names)
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = Sheet_Names(i)
    End With
  Next i


' SUB - Imports all the data to the correct sheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TODO - Loop and import the data to the correct sheet


For i In UBound(Sheet_Names)
  Sheets(Sheet_Names(i)).Select

  With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
      "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Y:\Data Intelligence\Code_Database\Data_Intelligence_Cod" _
      , _
      "e_Database.accdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:" _
      , _
      "Database Password=BORIS;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:" _
      , _
      "Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=" _
      , _
      "False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:" _
      , _
      "Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass Choice" _
      , "Field Validation=False"), Destination:=Range("$A$1")).QueryTable
      .CommandType = xlCmdTable
      .CommandText = Array(Access_Tbl_Names(i))
      .SourceDataFile = Access_File_Path
      .ListObject.DisplayName = Sheet_Names(i)
      .Refresh BackgroundQuery:=False
  End With

  ' Breaks the link so the database isn't locked in read only.
  Sheets(Sheet_Names(i)).ListObjects(Sheet_Names(i)).Unlink

Next i


End Sub
