MssqlCsvTransfer
===


### 快速使用
上傳CSV檔案
```ps1
irm raw.githubusercontent.com/hunandy14/MssqlCsvTransfer/master/Import-MssqlCsv.ps1|iex
Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -CsvPath "csv\Data.csv" | Out-Null
```

下載CSV檔案
```ps1
irm raw.githubusercontent.com/hunandy14/MssqlCsvTransfer/master/Export-MssqlCsv.ps1|iex
Export-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -OpenOutDir | Out-Null
```
