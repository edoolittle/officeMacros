' Attribute VB_Name = "SaveAttachments"
' https://deepsonline.com/blogs/vba-macro-to-save-attachments-multiple-selected-emails-outlook/
' Requires GetOutputDirectory() function

' Requires reference to Microsoft Scripting Runtime (SCRRUN.DLL)
' Go to Tools -> References -> check Microsoft Scripting Runtime
' There may still be a problem; if so, see
' See https://softwaretested.com/files/what-is-scrrun-dll/
' Restart Outlook after following instructions re: regsvr32.exe

Public Sub SaveAttachments()

  'Note, this assumes you are in the a folder with e-mail messages when you run it.
  'It does not have to be the inbox, simply any folder with e-mail messages
  
  Dim App As New Outlook.Application
  Dim Exp As Outlook.Explorer
  Dim Sel As Outlook.Selection
  
  Dim AttachmentCnt As Integer
  Dim AttTotal As Integer
  Dim MsgTotal As Integer
  
  Dim outputDir As String
  Dim outputFile As String
  Dim fileExists As Boolean
  Dim cnt As Integer
  Dim strSubject As String

  Dim fso As FileSystemObject
    
  Set Exp = App.ActiveExplorer
  Set Sel = Exp.Selection
  Set fso = New FileSystemObject

  outputDir = GetOutputDirectory()
  If outputDir = "" Then
    MsgBox "You must pick an directory to save your files to. Exiting SaveAttachments.", vbCritical, "SaveAttachments"
    Exit Sub
  End If
    
  'Loop thru each selected item in the inbox
  For cnt = 1 To Sel.Count
    'If the e-mail has attachments...
    If Sel.Item(cnt).Attachments.Count > 0 Then
      MsgTotal = MsgTotal + 1
      'For each attachment on the message...
      For AttachmentCnt = 1 To Sel.Item(cnt).Attachments.Count
        'Get the attachment
        Dim att As Attachment
        Set att = Sel.Item(cnt).Attachments.Item(AttachmentCnt)
        outputFile = att.FileName
        'Forcing to give me option to choose file name
        Let strSubject = Sel.Item(cnt).SentOn & vbCrLf & Sel.Item(cnt).Subject & vbCrLf & "( From " & Sel.Item(cnt).SenderName & " )"
        outputFile = InputBox(strSubject & vbCrLf & vbCrLf & "Please enter a new name if needed, or hit cancel to skip this one file.give name cancel to exit", "File Name", outputFile)
        
        If outputFile = "" Then
            'Exit leaving fileexists true. That will be a flag not to write the file
            GoTo nextitem
         
         End If
         'Give an option to exit
         If outputFile = "cancel" Then
            GoTo earlyexit
         End If
        fileExists = fso.fileExists(outputDir + outputFile)
        Do While fileExists = True
          outputFile = InputBox("The file " + outputFile _
            + " already exists in the destination directory of " _
            + outputDir + ". Please enter a new name, or hit cancel to skip this one file.", "File Exists", outputFile)
          'If user hit cancel
          If outputFile = "" Then
            'Exit leaving fileexists true. That will be a flag not to write the file
            Exit Do
          End If
          fileExists = fso.fileExists(outputDir + outputFile)
        Loop
        
        'Save it to disk if the file does not exist
        If fileExists = False Then
          att.SaveAsFile (outputDir + outputFile)
          AttTotal = AttTotal + 1
        End If
nextitem:
      Next
    End If
  Next
earlyexit:
  'Clean up
  Set Sel = Nothing
  Set Exp = Nothing
  Set App = Nothing
  Set fso = Nothing
  
  'Let user know we are done
  Dim doneMsg As String
  doneMsg = "Completed saving " + Format$(AttTotal, "#,0") + " attachments in " + Format$(MsgTotal, "#,0") + " Messages."
  MsgBox doneMsg, vbOKOnly, "Save Attachments"
  
  Exit Sub
  
ErrorHandler:

  Dim errMsg As String
  errMsg = "An error has occurred. Error " + Err.Number + " " + Err.Description
  Dim errResult As VbMsgBoxResult
  errResult = MsgBox(errMsg, vbAbortRetryIgnore, "Error in Save Attachments")
  Select Case errResult
    Case vbAbort
      Exit Sub
      
    Case vbRetry
      Resume
      
    Case vbIgnore
      Resume Next
      
  End Select
    
End Sub

