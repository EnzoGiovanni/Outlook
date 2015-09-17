Const Directory As String = "your destination directory"
Sub SaveMail()
    Dim SelectedMsg As Outlook.Selection: Set SelectedMsg = Application.ActiveExplorer.Selection
    If SelectedMsg.Count > 0 Then
        Dim Message As MailItem
        For Each Message In SelectedMsg
            If Message.Class = olMail Then Message.SaveAs Directory & Message.ConversationTopic & ".msg", OlSaveAsType.olMSG
        Next Message: Set Message = Nothing
    End If
    Set SelectedMsg = Nothing
End Sub
