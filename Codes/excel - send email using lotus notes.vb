Sub Send_Email_via_Lotus_Notes()
    Dim Maildb As Object
    Dim MailDoc As Object
    Dim Body As Object
    Dim Session As Object

    'Start a session of Lotus Notes
        Set Session = CreateObject("Lotus.NotesSession")
    'This line prompts for password of current ID noted in Notes.INI
        Call Session.Initialize
    'or use below to provide password of the current ID (to avoid Password prompt)
        'Call Session.Initialize("<password>")
    'Open the Mail Database of your Lotus Notes
        Set Maildb = Session.GETDATABASE("", "C:\Program Files\Lotus\Notes\Data\mail\alvesfr.nsf")
        If Not Maildb.IsOpen = True Then Call Maildb.Open
    'Create the Mail Document
        Set MailDoc = Maildb.CREATEDOCUMENT
        Call MailDoc.ReplaceItemValue("Form", "Memo")
    'Set the Recipient of the mail
        Call MailDoc.ReplaceItemValue("SendTo", "EMAIL HERE")
    'Set subject of the mail
        Call MailDoc.ReplaceItemValue("Subject", "TESTE DE EMAIL")
    'Create and set the Body content of the mail
        Set Body = MailDoc.CREATERICHTEXTITEM("Body")
        Call Body.AppendText("TESTE DE EMAIL")
    'Example to create an attachment (optional)
        Call Body.AddNewLine(2)
        Call Body.EMBEDOBJECT(1454, "", "C:\Documents and Settings\alvesfr\Desktop\Ramais Uteis.xls", "Attachment")
    'Example to save the message (optional) in Sent items
        MailDoc.SAVEMESSAGEONSEND = True
    'Send the document
    'Gets the mail to appear in the Sent items folder
        Call MailDoc.ReplaceItemValue("PostedDate", Now())
        Call MailDoc.SEND(False)
    'Clean Up the Object variables - Recover memory
        Set Maildb = Nothing
        Set MailDoc = Nothing
        Set Body = Nothing
        Set Session = Nothing
End Sub