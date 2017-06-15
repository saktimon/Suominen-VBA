' Created by Sakari Timonen - Testify Oy 2017-06 for free usage of Suominen Oy
' This is a macro for copying defects from Test script to a Defect log automatically and works for predefined templates
' If Step status (Column-H) = Defect && Defect ID doesn't exists yet (Column-Q)
' Copies found defect rows to 1st empty rows on a predefined Defect log (excluding Defect log A-column for ID)
' And then copies the ID from Defect log to Test script

Sub CopyDefects()

Dim LastRow As Integer, i As Integer, erow As Long
Dim DefectLogURL As String, DefectLogName As String
Dim DefectLogWasOpened As Boolean
Dim DefectLogWB As Workbook
Dim TestScript As Workbook
DefectLogWasOpened = False


Set TestScript = ThisWorkbook


'Just for debuggin ContentTypeProperties
'For i = 1 To ThisWorkbook.ContentTypeProperties.Count
'Debug.Print (ThisWorkbook.ContentTypeProperties(i).Value)
'Next i

'DefectLogURL is fetched from Sharepoint attribute DefectLog
DefectLogURL = ThisWorkbook.ContentTypeProperties.Item("DefectLog").Value
Debug.Print ("Defect log URL: " & DefectLogURL)

'DefectLogURL = "https://testifyoy.sharepoint.com/Shared%20Documents/DEFECT%20log.xlsx"
DefectLogName = DigFilename(DefectLogURL)
Debug.Print ("Filename after splitting: " & DefectLogName)



' This public function found by googling parses URL to appropriate format
DefectLogURL = Parse_Resource(DefectLogURL)

'Debug.Print ("Defect log parsed:" & DefectLogURL)

' Finds how many rows in script has values, i.e. how many rows to check for defects
LastRow = Cells.Find(What:="*", _
                    after:=Range("A1"), _
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
    Set DefectLogWB = ActiveWorkbook
'    SetAttr DefectLogURL, vbNormal
    DefectLogWasOpened = True
    Debug.Print ("Defect log was opened")
Else
    Set DefectLogWB = Workbooks(DefectLogName)
    Debug.Print ("Defect log was already open")
End If

'Find the 1st empty row from Defect log (ID in Column A is disregarded
erow = FirstBlankRow(DefectLogWB.Worksheets("Defect log").Range("B26:K426"))

'Copy defect row to Defect log
'ThisWorkbook.Activate
TestScript.Sheets(2).Range(Cells(i, 1), Cells(i, 16)).Copy
DefectLogWB.Worksheets("Defect log").Cells(erow, 3).PasteSpecial xlPasteValues

'Copy test script name to Defect log
DefectLogWB.Worksheets("Defect log").Cells(erow, 2) = ThisWorkbook.Name

DefectLogWB.Save

'Copy-paste defect ID to test script
DefectLogWB.Worksheets("Defect log").Cells(erow, 1).Copy ThisWorkbook.ActiveSheet.Cells(i, 17)

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

Function Parse_Resource(URL As String)
 'Uncomment the below line to test locally without calling the function & remove argument above
 'Dim URL As String
 Dim SplitURL() As String
 Dim i As Integer
 Dim WebDAVURI As String


 'Check for a double forward slash in the resource path. This will indicate a URL
 If Not InStr(1, URL, "//", vbBinaryCompare) = 0 Then

     'Split the URL into an array so it can be analyzed & reused
     SplitURL = Split(URL, "/", , vbBinaryCompare)

     'URL has been found so prep the WebDAVURI string
     WebDAVURI = "\\"

     'Check if the URL is secure
     If SplitURL(0) = "https:" Then
         'The code iterates through the array excluding unneeded components of the URL
         For i = 0 To UBound(SplitURL)
             If Not SplitURL(i) = "" Then
                 Select Case i
                     Case 0
                         'Do nothing because we do not need the HTTPS element
                     Case 1
                         'Do nothing because this array slot is empty
                     Case 2
                     'This should be the root URL of the site. Add @ssl to the WebDAVURI
                         WebDAVURI = WebDAVURI & SplitURL(i) & "@ssl"
                     Case Else
                         'Append URI components and build string
                         WebDAVURI = WebDAVURI & "\" & SplitURL(i)
                 End Select
             End If
         Next i

     Else
     'URL is not secure
         For i = 0 To UBound(SplitURL)

            'The code iterates through the array excluding unneeded components of the URL
             If Not SplitURL(i) = "" Then
                 Select Case i
                     Case 0
                         'Do nothing because we do not need the HTTPS element
                     Case 1
                         'Do nothing because this array slot is empty
                         Case 2
                     'This should be the root URL of the site. Does not require an additional slash
                         WebDAVURI = WebDAVURI & SplitURL(i)
                     Case Else
                         'Append URI components and build string
                         WebDAVURI = WebDAVURI & "\" & SplitURL(i)
                 End Select
             End If
         Next i
     End If
  'Set the Parse_Resource value to WebDAVURI
  Parse_Resource = WebDAVURI
 Else
 'There was no double forward slash so return system path as is
     Parse_Resource = URL
 End If


 End Function
 
 Function DigFilename(URL As String)

 'Uncomment the below line to test locally without calling the function & remove argument above
 Dim SplitURL() As String

 'Split the URL into an array
 SplitURL = Split(URL, "/", , vbBinaryCompare)
 ' Debug.Print (SplitURL(UBound(SplitURL)))
 DigFilename = Replace(SplitURL(UBound(SplitURL)), "%20", " ")

End Function
