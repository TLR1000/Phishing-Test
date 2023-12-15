Sub amItested()
    'This macro tests for the appearance of phishing test mails in your inbox.

    Dim MessageInfoString As String
    Dim UserMessages As String ' Variable to store user messages

    ' Declare your variables
    Dim myNameSpace As Outlook.NameSpace
    Dim myInbox As Outlook.Folder
    Dim myitems As Outlook.Items
    Dim myDestFolder As Outlook.Folder
    Dim myitem As Object ' Change the data type to Object to handle all types of items

    Set myNameSpace = Application.GetNamespace("MAPI")
    Set myInbox = myNameSpace.GetDefaultFolder(olFolderInbox)
    Set myitems = myInbox.Items

    ' Set the destination folder
    counterProcessed = 0

    ' Define the X-PHIS attributes you want to check for
    Dim xPhisAttributes As Variant
    ' Use the correct case for the header names
    xPhisAttributes = Array("X-PHISH-CRID", "X-PHISHTEST") ' Add your desired attributes

    ' Loop through the emails in the Inbox
    If Not myitems Is Nothing Then
        For Each myitem In myitems
            ' Check if the item is a MailItem
            If TypeName(myitem) = "MailItem" Then
                ' Access the internet headers
                Dim internetHeaders As String
                ' Use the correct case for the property tag
                internetHeaders = myitem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
                
                ' Check if the X-PHIS attributes are present in the headers
                Dim attributeFound As Boolean
                attributeFound = False
                For Each xPhisAttribute In xPhisAttributes
                    ' Use the correct case for the header name
                    If InStr(1, internetHeaders, xPhisAttribute, vbTextCompare) > 0 Then
                        attributeFound = True
                        Exit For
                    End If
                Next xPhisAttribute
                
                ' If any of the specified attributes are found, move the email
                If attributeFound Then
                    ' Get email information
                    Dim emailSubject As String
                    Dim emailSender As String
                    emailSubject = myitem.Subject
                    emailSender = myitem.SenderEmailAddress
                    ' Build the user message and add it to UserMessages
                    UserMessages = UserMessages & "Look for email titled '" & emailSubject & "' from " & emailSender & vbCrLf
                    ' Move the email (uncomment this line if needed)
                    'myitem.Move myDestFolder
                    counterProcessed = counterProcessed + 1
                End If
            End If
        Next myitem
    Else
        MsgBox "No items in the Inbox", vbInformation, "Done"
    End If

    ' Combine all user messages into one MsgBox
    If counterProcessed > 0 Then
        MessageInfoString = "Found " & counterProcessed & " email(s) to test you!" & vbCrLf & vbCrLf & UserMessages
    Else
        MessageInfoString = "No matching emails found."
    End If

    ' Display the combined message in a MsgBox
    MsgBox MessageInfoString, vbInformation, "Done"
End Sub
