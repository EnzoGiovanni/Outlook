Const Directory As String = "your destination directory"

Sub SaveMail()
    Dim SelectedMsg As Outlook.Selection: Set SelectedMsg = Application.ActiveExplorer.Selection
    If SelectedMsg.Count > 0 Then
        Dim Message As MailItem
        For Each Message In SelectedMsg
            If Message.Class = olMail Then Message.SaveAs Directory & CleanStringForFileName(Message.ConversationTopic) & ".msg", OlSaveAsType.olMSG
        Next Message: Set Message = Nothing
    End If
    Set SelectedMsg = Nothing
End Sub

Function CleanStringForFileName(Char As String) As String
    If Len(Char) > 0 Then
        Char = Replace(Char, "<", " ")
        Char = Replace(Char, ">", " ")
        Char = Replace(Char, ":", " ")
        Char = Replace(Char, Chr(34), " ")
        Char = Replace(Char, "/", "-")
        Char = Replace(Char, "|", " ")
        Char = Replace(Char, "?", " ")
        Char = Replace(Char, "*", " ")
        Char = Trim(Char)
        Do While InStr(1, Char, "  ", vbTextCompare)
            Char = Replace(Char, "  ", " ")
        Loop
        If Len(Char) <= 255 Then Char = Mid(Char, 1, 255)
        CleanStringForFileName = Char
    End If
End Function
