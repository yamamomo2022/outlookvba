Attribute VB_Name = "M_sendmail"
Option Explicit


Public Sub sendmail(Optional ByVal strSubject As String = "�o�Ε� ������")
    
    Dim Response As VbMsgBoxResult
    Dim objMail As MailItem
    Dim strRecipient As String
    Dim strCC As String
    Dim strBCC As String
    Dim strBody As String
    
    ' ���M�惁�[���A�h���X�ACC�ABCC�A�����A�{����ݒ�
    strRecipient = "example@example.com"  ' ���M��̃��[���A�h���X�������ɋL�����Ă��������B
    strCC = "cc@example.com"              ' CC�̃��[���A�h���X�������ɋL�����Ă��������B
    strBCC = "bcc@example.com"            ' BCC�̃��[���A�h���X�������ɋL�����Ă��������B
    strBody = ""
    
    Response = MsgBox("���[���𑗐M���܂����H", vbYesNo + vbQuestion, "���[�����M�̊m�F")
    
    
    
    If Response = vbYes Then
        Set objMail = Application.CreateItem(olMailItem)
        With objMail
            .To = strRecipient
            .CC = strCC
            .BCC = strBCC
            .Subject = strSubject
            .Body = strBody
            .Send
        End With
    End If

End Sub


Private Sub Application_Startup()

    Application.ActiveExplorer.WindowState = olMinimized
    UserForm1.Show vbModeless


End Sub
