Attribute VB_Name = "Module1"
Option Explicit

Sub Create_Email()

'Defining outlook variables
Dim OutApp As Object
Dim OutMail As Object

'Allocated
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

Dim NextParagraph As String

NextParagraph = vbNewLine & vbNewLine

With OutMail
    .To = "info@excelspreadsheet-support.com"
    .cc = "info@excelspreadsheet-support.com"
    .bcc = "info@excelspreadsheet-support.com"
    .Subject = "Let us see if subject line"
    
    
    .Body = "Hi James" & NextParagraph & _
            "let me know ok?" & NextParagraph & _
            "Looks like this works!"
    
    .Attachments.Add ("C:\Users\Jon\OneDrive - jaytsystems\VBAPlayground.xlsm")
    .Attachments.Add ("C:\Users\Jon\OneDrive - jaytsystems\test.txt.txt")
    
    'Email Account already allocated into Outlook
    'Set .SendUsingAccount = OutApp.Session.Accounts.Item("machground@msn.com")
    
    'Send / Display / Save / From
    .Display
    
End With


End Sub
