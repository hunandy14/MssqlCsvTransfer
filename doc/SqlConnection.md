# SQL 連接方式說明

本文件說明 `Get-SqlQueryResult` 函數支援的各種 SQL Server 連接方式。

## 支援的連接方式

### 1. 連接字串 (Connection String)

直接使用連接字串是最簡單的方式。

```powershell
Get-SqlQueryResult -Connection '
    Data Source=UX533-PC;
    Initial Catalog=CHG;
    User ID=chg;
    Password=1230
' -Query 'SELECT * FROM [CHG].[CHG].[Table02]'
```

### 2. 雜湊表 (Hashtable) 

使用雜湊表可以更結構化地指定連接參數。

```powershell
Get-SqlQueryResult -Connection @{
    DataSource     = 'UX533-PC'
    InitialCatalog = 'CHG'
    UserID         = 'chg'
    Password       = '1230'
} -Query 'SELECT * FROM [CHG].[CHG].[Table02]'
```

### 3. SqlConnectionStringBuilder 物件

使用 `SqlConnectionStringBuilder` 可以更安全地建立連接字串。

```powershell
Get-SqlQueryResult -Connection ([Data.SqlClient.SqlConnectionStringBuilder]@{
    DataSource     = 'UX533-PC'
    InitialCatalog = 'CHG'
    UserID         = 'chg'
    Password       = '1230'
}) -Query 'SELECT * FROM [CHG].[CHG].[Table02]'
```

### 4. SqlConnection 物件

直接使用 `SqlConnection` 物件進行連接。

```powershell
# 建立 SqlConnection 物件
$conn = ([Data.SqlClient.SqlConnection](
    ([Data.SqlClient.SqlConnectionStringBuilder]@{
        DataSource     = 'UX533-PC'
        InitialCatalog = 'CHG'
        UserID         = 'chg'
        Password       = '1230'
    }).ConnectionString
))

# 建立連接
Get-SqlQueryResult -Connection $conn -Query 'SELECT * FROM [CHG].[CHG].[Table02]'

# 釋放資源
$conn.Close()
$conn.Dispose()
```

> 由外部傳入的 `[Data.SqlClient.SqlConnection]` 物件記得自行釋放資源
