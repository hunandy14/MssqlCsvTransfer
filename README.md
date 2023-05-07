MssqlCsvTransfer
===


### 快速使用
上傳CSV檔案
```ps1
irm bit.ly/ImpMssql|iex; Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -CsvPath "csv\Data.csv"
```

下載CSV檔案
```ps1
irm bit.ly/ExpMssql|iex; Export-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]"
```
