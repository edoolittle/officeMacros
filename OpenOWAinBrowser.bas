' VBA code for Outlook Classic to open Outlook for Web by pressing a button
' To install:
' - enable Developer tab in Outlook
' - Click on Developer, Visual Basic
' - right click on Project1, then Insert > Module
' - In the new module that is created, paste the code in this file
' - Save, and close if desired
' - Back in Outlook, execute the code by clicking Developer, Macros, Project1.OpenOWAinBrowser
' - (You may have to adjust security settings to enable the execution of macros)
' - Outlook for the Web should open in a browser tab
' To make a quick access button:
' - Click on Customize Quick Access Toolbar, the down-arrow like symbol near the very top left of the Outlook window
' - Select More Commands
' - Select Choose Commands From: Macros
' - Select Project1.OpenOWAinBrowser, then Add->
' - Click Modify if you want to change the icon or the tool tip (to, say OWA instead of Project1.OpenOWAinBrowser)
' Done!

' Declare the ShellExecute API function
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

' Add OWA Button to open Outlook on the web
Sub OpenOWAinBrowser()
  Dim fullURL As String
  fullURL = "https://outlook.office.com/mail/"
  ShellExecute 0, "open", fullURL, vbNullString, vbNullString, 1
End Sub
