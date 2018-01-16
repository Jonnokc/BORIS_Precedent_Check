Option Private Module

Sub Data_Import()

 Dim sheet As Worksheet
 Dim Sheet_Names As Variant
 Dim Access_Tbl_Names As Variant
 Dim Access_File_Path As Variant
 Dim LastRow As Long
 Dim Data_Done As Single

 ' Enables settings to make data import functional
 Call Utils.Enable_Settings


 Sheet_Names = Array("Medications", "Keywords", "Valid_Code_Systems", "Previously_Mapped")

 Access_Tbl_Names = Array("Q_All_Medication_Raw_Code_Displays_Excel", "Q_Proprietary_Code_Display_Keywords", "Q_Mappable_Raw_Code_System_Names", "Q_Raw_Codes_To_Concept_Alias_2017")

 Access_File_Path = "Y:\Data Intelligence\Code_Submittion_Database\CodeFeedbackDatabase.accdb"



' SUB - Checks to confirm path to Access Database is mapped correctly. If not then it asks the user to get it.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ' Progress status display
 Call Utils.Progress("Checking Data Location....")

 ' Checks file path and confirms file location
 Set fso = CreateObject("scripting.filesystemobject")
 With fso
   Path_Checker = Len(Dir("Y:\Data Intelligence\Code_Database\Data_Intelligence_Code_Database.accdb")) <> 0
 End With

 If Path_Checker = 0 Then
   Access_Database_Check = Utils.GetOpenFile()

   If Access_Database_Check = False Then
     GoTo DataError
   Else
     Access_File_Path = Access_Database_Check
   End If
 End If

 Set fso = Nothing

 ' Closes the status popup
 Call Utils.Progress_Close


' SUB - Delets all the extra sheets and then readds them to make sure things remain clean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ' Progress status display
 Call Utils.Progress("Cleaning Sheets")

 ' Disables confirm prompt so sheets are deleted without user confirming
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
 ' re-enables security prompt
 Application.DisplayAlerts = True


 ' Creates the sheets within the workbook
 For i = 0 To UBound(Sheet_Names)
   With ThisWorkbook
     .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = Sheet_Names(i)
   End With
 Next i

' Closes the status popup
 Call Utils.Progress_Close


' SUB - Imports all the data to the correct sheet
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ufProgress.LabelProgress.Width = 0
 ufProgress.Show
 For i = 0 To UBound(Sheet_Names)

   Data_Done = i / UBound(Sheet_Names)

   ' updates progress bar
   With ufProgress
     .LabelCaption = "Getting All the Data! On dataset " & (i + 1) & " of " & (UBound(Sheet_Names) + 1)
     .LabelProgress.Width = Data_Done * (.FrameProgress.Width)
   End With
   DoEvents

   Sheets(Sheet_Names(i)).Select

   With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
         "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=Y:\Data Intelligence\Code_Database\Data_Intelligence_Cod" _
         , _
         "e_Database.accdb;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:D" _
         , _
         "atabase Password="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Glob" _
         , _
         "al Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=Fals" _
         , _
         "e;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Supp" _
         , _
         "ort Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceFiel" _
         , "d Validation=False"), Destination:=Range("$A$1")).QueryTable
     .CommandType = xlCmdTable
     .CommandText = Array(Access_Tbl_Names(i))
     .BackgroundQuery = True
     .RefreshStyle = xlInsertDeleteCells
     .SaveData = True
     .AdjustColumnWidth = True
     .PreserveColumnInfo = True
     .SourceDataFile = Access_File_Path
     .ListObject.DisplayName = Sheet_Names(i)
     .Refresh BackgroundQuery:=False
   End With

   ' Breaks the link so the database isn't locked in read only.
   Sheets(Sheet_Names(i)).ListObjects(Sheet_Names(i)).Unlink

   ' Closes progress bar when done
   If i = UBound(Sheet_Names) Then
     Unload ufProgress
   End If

 Next i


  ' Removes duplicates from the Previously_Mapped sheet so there is only one Raw Code Display
  With Sheets(Sheet_Names(3))
    Sheets(Sheet_Names(3)).Select
    Set StartCell = .Range("A1")
    LastRow = StartCell.SpecialCells(xlCellTypeLastCell).row
    .Range(Sheet_Names(3)).RemoveDuplicates Columns:=1, Header:=xlYes
  End With

' Disables settings to improve performance
Call Utils.Disable_Settings

Exit Sub


DataError:
' Turns back on settings so Excel is responsive.
Call Utils.Enable_Settings

MsgBox ("Something went wrong while trying to connect to the data. Exiting now. Please re-run and confirm you pick the correct data source.")


End Sub
