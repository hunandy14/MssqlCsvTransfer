MssqlCsvTransfer
===

## 快速使用

下載CSV檔案

```ps1
irm raw.githubusercontent.com/hunandy14/MssqlCsvTransfer/refs/heads/dev/2.0/Export-MssqlToCsv.ps1|iex; {
  # 連接資訊
  $serverInstance = "UX533-PC"
  $database = "CHG"
  $username = "chg"
  $password = "1230"
  
  # 建立連接字串
  $connectionString = "Server=$serverInstance;Database=$database;User Id=$username;Password=$password"
  
  # 查詢字串
  $query = "SELECT * FROM [CHG].[CHG].[Table02]"
  
  # 執行查詢
  $data = Get-SqlQueryResult -ConnectionString $connectionString -Query $query -Verbose -Raw
  
  # 輸出到CSV
  $data | ConvertTo-CsvString | Set-Content "tmp\CHG.CHG.Table02.csv"
}
```

上傳CSV檔案

```ps1

```




<br><br><br>

### Export-MssqlCsv

```ps1

```

> 輸出的CSV是根據SSMS標準輸出的
> 不過這標準官方沒寫我自己測的詳細可以參考 [test-csv-output-data](https://github.com/hunandy14/MssqlCsvTransfer/tree/dev/2.0/test-csv-output-data)



<br><br><br>

### Import-MssqlCsv

```ps1

```
