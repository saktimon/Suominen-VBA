Attribute VB_Name = "Module1"
' Created by Sakari Timonen - Testify Oy 2017-06 for free usage of Suominen Oy
' This is a macro for copying defects from Test script to a Defect log automatically and works for predefined templates
' If Step status (Column-H) = Defect && Defect ID doesn't exists yet (Column-Q)
' Copies found defect rows to 1st empty rows on a predefined Defect log (excluding Defect log A-column for ID)
' And then copies the ID from Defect log to Test script

Sub CopyDefects()

Dim LastRow As Integer, i As Integer, erow As Long
Dim DefectLogURL As String, DefectLogName As String
Dim DefectLogWasOpened As Boolean
Dim DefectLog As Workbook
Dim TestScript As Workbook
DefectLogWasOpened = False

Set TestScript = ThisWorkbook

' Remember to update the Defect log path here on all scripts!
DefectLogURL = "https://testifyoy.sharepoint.com/Shared%20Documents/DEFECT%20log.xlsx"
DefectLogName = "DEFECT log.xlsx"

'Debug.Print ("Defect log URL: " & DefectLogURL)

' This public function found by googling parses URL to appropriate format
DefectLogURL = Parse_Resource(DefectLogURL)

'Debug.Print ("Defect log parsed:" & DefectLogURL)

' Finds how many rows in script has values, i.e. how many rows to check for defects
LastRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
                    
Debug.Print ("Rows to check: " & LastRow)

' Goes through each row on test script
For i = 2 To LastRow

'If row has a defect and doesn't have Defect ID yet, it is selected for copying
If (Cells(i, 8) = "Defect" And Cells(i, 17) = "") Then
Debug.Print ("Defect row: " & i)

' Old Copying
'Range(Cells(i, 1), Cells(i, 16)).Select
'Selection.Copy

' Defect log file is opened if not already so
If Not IsWorkBookOpen(DefectLogName) Then
    Workbooks.Open Filename:=DefectLogURL, ReadOnly:=False, Notify:=False
    ActiveWorkbook.LockServerFile
    Set DefectLog = ActiveWorkbook
'    SetAttr DefectLogURL, vbNormal
    DefectLogWasOpened = True
    Debug.Print ("Defect log was opened")
Else
    Set DefectLog = Workbooks(DefectLogName)
    Debug.Print ("Defect log was already open")
End If

'Find the 1st empty row from Defect log (ID in Column A is disregarded
erow = FirstBlankRow(DefectLog.Worksheets("Defect log").Range("B26:K426"))

'Copy defect row to Defect log
'ThisWorkbook.Activate
TestScript.Sheets(2).Range(Cells(i, 1), Cells(i, 16)).Copy
DefectLog.Worksheets("Defect log").Cells(erow, 3).PasteSpecial xlPasteValues

'Copy test script name to Defect log
DefectLog.Worksheets("Defect log").Cells(erow, 2) = ThisWorkbook.Name

DefectLog.Save

'Copy-paste defect ID to test script
DefectLog.Worksheets("Defect log").Cells(erow, 1).Copy ThisWorkbook.ActiveSheet.Cells(i, 17)

TestScript.Save
Debug.Print ("Defect copied successfully: " & Cells(i, 17))

End If

Next i

If DefectLogWasOpened Then
    Workbooks(DefectLogName).Close
    Debug.Print ("Defect log Closed")
End If

End Sub


Function FirstBlankRow(ByVal rngToSearch As Range) As Long
   Dim R As Range
   Dim C As Range
   Dim RowIsBlank As Boolean

   For Each R In rngToSearch.Rows
      RowIsBlank = True
      For Each C In R.Cells
         If IsEmpty(C.Value) = False Then RowIsBlank = False
      Next C
      If RowIsBlank Then
      
         FirstBlankRow = R.Row
         Debug.Print ("Empty row: " & R.Row)
         If RowIsBlank Then Exit For
      End If
   Next R
   
End Function

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function



