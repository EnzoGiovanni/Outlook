Const Directory As String = "your destination directory"
Sub SaveMail()
    Dim SelectedMsg As Outlook.Selection: Set SelectedMsg = Application.ActiveExplorer.Selection
    If SelectedMsg.Count > 0 Then
        Dim Message As MailItem
        For Each Message In SelectedMsg.MailItem
            Trash = Message.SaveAs(Directory & SelectedMsg(elt).ConversationTopic & ".msg", OlSaveAsType.olMSG)
        Next Message
    End If
End Sub
