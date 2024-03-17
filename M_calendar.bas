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
    
    ' Outlookアプリケーションを取得
    Set olApp = New Outlook.Application
    ' Namespaceを取得
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' カレンダーフォルダーを取得 (デフォルトの予定表)
    Set olFolder = olNamespace.GetDefaultFolder(olFolderCalendar)
    
    ' JSON用のテキスト作成
    strJSON = "{""events"":["
       
    
'    ' 現在の日付から1か月後までの期間を設定
'    startDate = Date
'    endDate = DateAdd("m", -1, startDate)
    
    ' 現在の日付から1か月後までの期間を設定
    endDate = Date
    startDate = DateAdd("m", -1, startDate)
    
    ' カレンダー内の予定を走査
    For Each olApt In olFolder.Items
        If TypeOf olApt Is AppointmentItem Then
            ' 予定の開始日時が指定した期間内にあるか確認
            If olApt.Start >= startDate And olApt.Start <= endDate Then
                ' イベントの情報をJSONに追加
                strJSON = strJSON & "{""Subject"":""" & olApt.Subject & ""","
                strJSON = strJSON & """Start"":""" & Format(olApt.Start, "yyyy-mm-ddThh:mm:ss") & ""","
                strJSON = strJSON & """End"":""" & Format(olApt.End, "yyyy-mm-ddThh:mm:ss") & """},"
            End If
        End If
    Next olApt
    
    ' 不要な末尾のカンマを削除
    strJSON = Left(strJSON, Len(strJSON) - 1)
    
    ' JSONテキストを完成させる
    strJSON = strJSON & "]}"
    
    ' JSONデータをファイルに書き出し
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile("C:\Users\viole\Desktop\ExportedCalendarEvents.json", True)
    objFile.Write strJSON
    objFile.Close
    
    ' クリーンアップ
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
    ' イベントの情報をJSONに追加
    strJSON = strJSON & "{""Subject"":""" & olApt.Subject & ""","
    strJSON = strJSON & """Start"":""" & olApt.Start & ""","
    strJSON = strJSON & """End"":""" & olApt.End & """},"
End Function

