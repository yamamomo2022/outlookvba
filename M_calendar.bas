Attribute VB_Name = "M_calendar"
Option Explicit


Public Sub ExportCalendarEventsToJSON()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olApt As Outlook.AppointmentItem
    Dim objFSO As Object
    Dim objFile As Object
    Dim strJSON As String
    Dim i As Integer
    Dim startDate As Date
    Dim endDate As Date
    
    ' Outlook�A�v���P�[�V�������擾
    Set olApp = New Outlook.Application
    ' Namespace���擾
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' �J�����_�[�t�H���_�[���擾 (�f�t�H���g�̗\��\)
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' JSON�p�̃e�L�X�g�쐬
    strJSON = "{""events"":["
       
    
'    ' ���݂̓��t����1������܂ł̊��Ԃ�ݒ�
'    startDate = Date
'    endDate = DateAdd("m", -1, startDate)
    
    ' ���݂̓��t����1������܂ł̊��Ԃ�ݒ�
    endDate = Date
    startDate = DateAdd("m", -1, startDate)
    
    ' �J�����_�[���̗\��𑖍�
    For Each olApt In olFolder.Items
        If TypeOf olApt Is AppointmentItem Then
            ' �\��̊J�n�������w�肵�����ԓ��ɂ��邩�m�F
            If olApt.Start >= startDate And olApt.Start <= endDate Then
                ' �C�x���g�̏���JSON�ɒǉ�
                strJSON = strJSON & "{""Subject"":""" & olApt.Subject & ""","
                strJSON = strJSON & """Start"":""" & Format(olApt.Start, "yyyy-mm-ddThh:mm:ss") & ""","
                strJSON = strJSON & """End"":""" & Format(olApt.End, "yyyy-mm-ddThh:mm:ss") & """},"
            End If
        End If
    Next olApt
    
    ' �s�v�Ȗ����̃J���}���폜
    strJSON = Left(strJSON, Len(strJSON) - 1)
    
    ' JSON�e�L�X�g������������
    strJSON = strJSON & "]}"
    
    ' JSON�f�[�^���t�@�C���ɏ����o��
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile("C:\Users\viole\Desktop\ExportedCalendarEvents.json", True)
    objFile.Write strJSON
    objFile.Close
    
    ' �N���[���A�b�v
    Set objFile = Nothing
    Set objFSO = Nothing
    Set olApt = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
    
'    MsgBox "Calendar events exported to JSON successfully!", vbInformation
    Debug.Print "Calendar events exported to JSON successfully!", vbInformation
End Sub

Private Function calendar2json() As String
    ' �C�x���g�̏���JSON�ɒǉ�
    strJSON = strJSON & "{""Subject"":""" & olApt.Subject & ""","
    strJSON = strJSON & """Start"":""" & olApt.Start & ""","
    strJSON = strJSON & """End"":""" & olApt.End & """},"
End Function

