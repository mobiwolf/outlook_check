Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

If TypeName(Item) <> "MailItem" Then Exit Sub

Dim intRet As Integer

'CHECK FOR BLANK SUBJECT LINE
If Item.Subject = "" Then
    intRet = MsgBox("警告:您的邮件缺少主题,请注意填写" & vbNewLine, vbOKOnly + vbMsgBoxSetForeground + vbExclamation, "缺少主题")
    If intRet = vbOK Then
        Cancel = True
        Exit Sub
    End If
End If

'CHECK FOR FORGETTING ATTACHMENT
Dim intRes As Integer
Dim strMsg As String
Dim strThismsg As String
Dim intOldmsgstart As Integer
Dim bForceAttch As Boolean

' Does not search for "Attach", but for all strings in an array that is defined here
Dim sSearchStrings(2) As String
Dim bFoundSearchstring As Boolean
Dim i As Integer

bForceAttch = True
bFoundSearchstring = False
sSearchStrings(0) = "attach"
sSearchStrings(1) = "enclose"
sSearchStrings(2) = "附件"

' intOldmsgstart = InStr(Item.Body, "-----Original Message-----")
intOldmsgstart = InStr(Item.Body, "发件人:")

If intOldmsgstart = 0 Then
    strThismsg = Item.Body + " " + Item.Subject
Else
    strThismsg = Left(Item.Body, intOldmsgstart) + " " + Item.Subject
End If

' The above if/then/else will set strThismsg to be the text of this message only,excluding old/fwd/re msg
' if the original included message is mentioning an attachment, ignore that Also includes the subject line at the end of the strThismsg string

For i = LBound(sSearchStrings) To UBound(sSearchStrings)
    If InStr(LCase(strThismsg), sSearchStrings(i)) > 0 Then
        bFoundSearchstring = True
        Exit For
    End If
Next i


If bFoundSearchstring Then
    If Item.Attachments.Count = 0 Then
        If bForceAttch Then
            intRet = MsgBox("警告:您的邮件缺少附件,请注意添加附件！！！" & vbNewLine, vbOKOnly + vbMsgBoxSetForeground + vbExclamation, "缺少附件")
            If intRet = vbOK Then
                Cancel = True
                Exit Sub
            End If
        Else
            strMsg = "警告:您的邮件缺少附件,请注意添加！！！" & vbNewLine & "确认是否发送?"
            intRet = MsgBox(strMsg, vbYesNo + vbMsgBoxSetForeground + vbDefaultButton2 + vbExclamation, "缺少附件")
            If intRet = vbNo Then
                Cancel = True
                Exit Sub
            End If
        End If
    End If
End If

End Sub