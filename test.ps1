$json = @"
{
  "events": [
    {
      "Subject": "test_test_test_test1",
      "Start": "2024-03-13T10:00:00",
      "End": "2024-03-13T12:00:00"
    },
    {
      "Subject": "test_test_test_test2",
      "Start": "2024-03-13T13:00:00",
      "End": "2024-03-13T15:00:00"
    },
    {
      "Subject": "test_test_test_test1",
      "Start": "2024-03-13T13:00:00",
      "End": "2024-03-13T15:00:00"
    },
    {
      "Subject": "test_test_test_test2",
      "Start": "2024-03-13T13:00:00",
      "End": "2024-03-13T15:00:00"
    }
  ]
}
"@

# JSON を PowerShell のオブジェクトに変換
$object = ConvertFrom-Json $json

# "Subject" によってイベントをグループ化
$groupedEvents = $object.events |  Group-Object -Property "Subject"

$personlist = @("person0","person1","person2")
foreach ($person in $personlist){
  write-host $person
}

# グループごとにイベントを表示
foreach ($group in $groupedEvents) {
    Write-Host "Subject: $($group.Name -split '_')"
    foreach ($event in $group.Group) {
        Write-Host "Start: $($event.Start), End: $($event.End)"
    }
    Write-Host
}



