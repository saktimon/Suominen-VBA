Attribute VB_Name = "Module2"
 Public Function Parse_Resource(URL As String)
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
