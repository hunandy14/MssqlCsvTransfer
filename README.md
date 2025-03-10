MssqlCsvTransfer
===

## 快速使用

下載CSV檔案

```ps1
irm raw.githubusercontent.com/hunandy14/MssqlCsvTransfer/refs/heads/dev/2.0/Export-MssqlToCsv.ps1|iex
Export-SqlServerTableToCsv @{
  DataSource     = 'UX533-PC'
  InitialCatalog = 'CHG'
  UserID         = 'chg'
  Password       = '1230'
} -TableName "[CHG].[CHG].[Table02]"
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
