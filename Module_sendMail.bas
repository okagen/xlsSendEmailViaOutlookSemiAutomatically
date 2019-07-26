Attribute VB_Name = "Module_sendMail"
Option Explicit

Enum shMail
    sTitle_Row = 2
    sTitle_col = 1
    sSignature_col = 2
    sBody_col = 3
    
    sList_row = 5
    sAddress_col = 1
    sCC_col = 2
    sToName_col = 3
    sAttachment = 4
End Enum

'----------------------------------------------------------------------
' Send e-mail main.
'----------------------------------------------------------------------
Function sendMail_main(ByVal ws As Worksheet)
    'Create Outlook object
    Dim objOutlook As Outlook.Application
    Set objOutlook = New Outlook.Application
    
    With ws
        'Get text data
        Dim sendTitle As String, sendSig As String, sendBody As String
        sendTitle = .Cells(shMail.sTitle_Row, shMail.sTitle_col).Value
        sendSig = .Cells(shMail.sTitle_Row, shMail.sSignature_col).Value
        sendBody = .Cells(shMail.sTitle_Row, shMail.sBody_col).Value

       Dim countList As Long
       For countList = shMail.sList_row To .Cells(shMail.sList_row, shMail.sAddress_col).End(xlDown).Row
       
            'Create mail item objects.
            Dim objMailItem As Outlook.MailItem
            Dim objAttach As Outlook.Attachments
            Set objMailItem = objOutlook.CreateItem(olMailItem)
            Set objAttach = objMailItem.Attachments
            
            'Get text data.
            Dim sendAddress As String, sendCC As String, sendName As String, sendAttach As String
            sendAddress = .Cells(countList, shMail.sAddress_col).Value
            sendCC = .Cells(countList, shMail.sCC_col).Value
            sendName = .Cells(countList, shMail.sToName_col).Value
            sendAttach = ThisWorkbook.Path & .Cells(countList, shMail.sAttachment).Value
             
            'Create the body of mail.
            Dim mailBody As String
            mailBody = createMailBody(sendBody, sendName, sendSig)

            With objMailItem
                .To = sendAddress
                .CC = sendCC
                .Subject = sendTitle
                .body = mailBody
            End With
           
            objAttach.Add sendAttach
            Set objAttach = Nothing
            
            'Preview e-mail.
            objMailItem.Display
            'objMailItem.Send
            Set objMailItem = Nothing
       Next countList
    End With
End Function

'----------------------------------------------------------------------
' Create the body of the e-mail
'----------------------------------------------------------------------
Function createMailBody(ByVal body As String, ByVal sendName As String, ByVal sendSig As String) As String
    Dim mBody As String
    mBody = Replace(body, "[Name of the recipient]", sendName)
    mBody = mBody & Chr(13) & Chr(13) & Chr(13) & sendSig
    createMailBody = mBody
End Function
'


