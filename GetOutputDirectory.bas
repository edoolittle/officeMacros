Attribute VB_Name = "GetOutputDirectory"
' https://www.deepsonline.com/blogs/vba-macro-to-save-attachments-multiple-selected-emails-outlook/

Public Function GetOutputDirectory() As String
 
  Dim retval As String 'Return Value
  
  Dim sMsg As String
  Dim cBits As Integer
  Dim xRoot As Integer
  
  Dim oShell As Object
  Set oShell = CreateObject("shell.application")

  sMsg = "Select a Folder To Output The Attachments To"
  cBits = 1
  xRoot = 17
  
  On Error Resume Next
      Dim oBFF
      Set oBFF = oShell.BrowseForFolder(0, sMsg, cBits, xRoot)
      If Err Then
        Err.Clear
        GetOutputDirectory = ""
        Exit Function
      End If
  On Error GoTo 0
  
  If Not IsObject(oBFF) Then
    GetOutputDirectory = ""
    Exit Function
  End If
  
  If Not (LCase(Left(Trim(TypeName(oBFF)), 6)) = "folder") Then
    retval = ""
  Else
    retval = oBFF.Self.Path
    
    'Make sure there's a \ on the end
    If Right(retval, 1) <> "\" Then
      retval = retval + "\"
    End If
  End If
  
  GetOutputDirectory = retval
  
End Function
