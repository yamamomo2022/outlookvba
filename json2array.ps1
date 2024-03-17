# JSONファイルを読み込む
$jsonFilePath = "ExportedCalendarEvents.json"
$jsonData = Get-Content $jsonFilePath -Raw | ConvertFrom-Json

# Excelオブジェクトを作成
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

$STARTROWINDEX = 5
$STARTCOLINDEX = 5

# 開始日と終了日を定義
$StartDate = Get-Date "2024-01-16"
$endDate = Get-Date "2024-02-15"

# 配列を初期化
$dateArray = @()
$workhourArray = @()

# 日にちの数値を順番に格納した配列を作成
while ($startDate -le $endDate) {
    $dateArray += $startDate.Day
    $workhourArray += @("")
    $startDate = $startDate.AddDays(1)
}

$header = @("Responsible_person","Client","Name","Duties") + $dateArray
for ($colIndex = 1; $colIndex -le $header.Count+1; $colIndex++) { 
    $worksheet.Cells.Item($STARTROWINDEX, $colIndex) = $header[$colIndex-1]
}

$postSubject = @()
$outputarray = @()
# イベントをExcelに書き込む
$colIndex = $STARTCOLINDEX 
$rowIndex = $STARTROWINDEX + 1
foreach ($event in $jsonData.events) {
    $date = [datetime]::Parse($event.Start)

    # 開始日時と終了日時の差を計算して時間差を取得
    $timeDifference = ([datetime]::Parse($event.End) - [datetime]::Parse($event.Start)).TotalHours

    
    # #初回以外の処理
    # if($outputarray.count -gt 0){
    #     $postSubjectList = $postSubject.Subject  -split '_'
    #     if 
    # } 

    if($postSubject -eq $event.Subject){
        $rowIndex--
    }

    for ($i = 0; $i -le $dateArray.count-1; $i++){
        # write-host $date.day.ToString(),$dateArray[$i].ToString()
        if ( $date.day -eq $dateArray[$i].ToString() ){
            $colIndex = $i + $STARTCOLINDEX
            if ($workhourArray[$i] -as [int]) {
                # 数字である場合の処理
                $workhourArray[$i] += $timeDifference
            } else {
                # 数字でない場合の処理
                $workhourArray[$i] = $timeDifference
            }
            break
        }
    }

    #Excelに記入
    $SubjectList = $event.Subject  -split '_'
    $outputarray = $SubjectList + $workhourArray
    for ($j = 1; $j -le $outputarray.count;$j++){
        $worksheet.Cells.Item($rowIndex, $j) = $outputarray[$j-1]
    }  


    $colIndex++
    $rowIndex++

    $postSubject = $event.Subject

}


# 列の幅を自動調整
$worksheet.Columns.Item(1).AutoFit() | Out-Null
$worksheet.Columns.Item(2).AutoFit() | Out-Null
$worksheet.Columns.Item(3).AutoFit() | Out-Null
$worksheet.Columns.Item(4).AutoFit() | Out-Null

# Excelファイルを保存
$excelFile = "C:\Users\viole\Desktop\ExportedCalendarEvents_test.xlsx"
$workbook.SaveAs($excelFile)

# Excelを閉じる
$excel.Quit()

Write-Host "finish"



