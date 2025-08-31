Attribute VB_Name = "ConvertEmailToEvent"
' https://www.datanumen.com/blogs/2-easy-ways-quickly-convert-email-meeting-invitation-outlook/
Sub ConvertEmailToEvent()
    Dim objMail As Outlook.MailItem
    Dim objMeeting As Outlook.AppointmentItem
    Dim objAttachment As Outlook.Attachment
    Dim strFilePath As String
 
    Select Case Outlook.Application.ActiveWindow.Class
           Case olInspector
                Set objMail = ActiveInspector.CurrentItem
           Case olExplorer
                Set objMail = ActiveExplorer.Selection.Item(1)
    End Select
 
    Set objMeeting = Outlook.Application.CreateItem(olAppointmentItem)
 
    'Copy Attachments
    If objMail.Attachments.Count > 0 Then
       For Each objAttachment In objMail.Attachments
           strFilePath = CStr(Environ("USERPROFILE")) & "\AppData\Local\Temp\" & objAttachment.FileName
           objAttachment.SaveAsFile (strFilePath)
 
           objMeeting.Attachments.Add strFilePath
 
           Kill strFilePath
       Next
    End If
 
    With objMeeting
        .Subject = objMail.Subject
        'Exclude Original Email Header
        .Body = objMail.Body
        .MeetingStatus = olNonMeeting
        .Display
    End With
End Sub

