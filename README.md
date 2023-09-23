MssqlCsvTransfer
===

## 快速使用
上傳CSV檔案
```ps1
irm bit.ly/ImpMssql|iex; Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8
```

下載CSV檔案
```ps1
irm bit.ly/ExpMssql|iex; Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -UTF8
```



<br><br><br>

### Import-MssqlCsv

```ps1
# 載入函式
irm bit.ly/ImpMssql|iex

# 上傳CSV檔案
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.csv"
# 上傳CSV檔案 (使用UTF8編碼) (預設是依據電腦語言決定)
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8
# 上傳CSV檔案 (移除原有表格內容)
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -CleanTable -UTF8

# 上傳CSV檔案 (檔案包含檔頭)
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.data" -Csv_RemoveQuotesHeaders -UTF8
# 上傳CSV檔案 (檔案包含雙引號)
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.data" -Csv_RemoveQuotes -UTF8
# 上傳CSV檔案 (檔案包數值包含逗號)
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.data" -Csv_ReplaceDelimiter "|`,|" -UTF8
# 上傳CSV檔案 (指定暫存檔案儲存位置) (使用 '.tmp' 會自動刪除)
Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" -Path "csv\Data.data" -Csv_RemoveQuotes -TempPath "data\Data.data" -UTF8

```

> 使用的是 BCP 直接將本地的 CSV 給傳到伺服器端  
> BCP 是微軟專門開發用來傳輸大檔用的程式，性能妥妥的有保障。  



<br><br><br>

### Export-MssqlCsv

```ps1
# 載入函式
irm bit.ly/ExpMssql|iex

# 下載表格存為CSV檔案 (自動以表格名命名)
Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[Employees]"
# 下載表格存為CSV檔案
Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[Employees]" -Path "csv\Employees.csv"
# 下載表格存為CSV檔案 (使用UTF8編碼) (預設是依據電腦語言決定)
Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[Employees]" -Path "csv\Employees.csv" -UTF8

# 下載自定義SqlQuerry (從檔案)
Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLPath "sql\Employees.sql" -Path "csv\Employees.csv" -UTF8
# 下載自定義SqlQuerry (從命令)
Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLQuery "Select * From CHG.CHG.Employees" -Path "csv\Employees.csv" -UTF8
"Select * From CHG.CHG.Employees" | Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Path "csv\Employees2.csv" -UTF8

# 下載表格存為CSV檔案 (儲存過程中實際用的Querry檔案) (使用 '.tmp' 會自動刪除)
Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[Employees]" -Path "csv\Employees.csv" -TempSqlPath "sql\Employees2.sql" -UTF8

```

> 使用的是 sqlcmd 直接將 sql 語句的結果輸出到檔案上  
> 為什麼不用 BCP 是因為他不支持外部sql檔案，用管道傳怕語句太長到時候出什麼奇怪的bug  
> 而為什麼需要外部 sql 檔案是因為輸出結果本身是不包含雙引號的，簡單的解法是在輸出的時候多過一層 sql 語句補上字段區隔的雙引號。  
