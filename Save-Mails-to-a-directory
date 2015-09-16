Const Directory As String = "your destination directory"
Sub SaveMail()
    Dim SelectedMsg As Outlook.Selection: Set SelectedMsg = Application.ActiveExplorer.Selection
    Dim elt As Long
    If SelectedMsg.Count > 0 Then
        For elt = 1 To SelectedMsg.Count Step 1
            Trash = SelectedMsg(elt).SaveAs(Directory & SelectedMsg(elt).ConversationTopic & ".msg", OlSaveAsType.olMSG)
        Next elt
    End If
End Sub
