$serverName   = "192.168.3.123"
$databaseName = "CHG"
$userName     = "kaede"
$password     = "1230"

$schemaName = "CHG"
$tableName = "Employees"
$Table = "[$databaseName].[$schemaName].[$tableName]"

# 建立连接到数据库的 SqlConnection 对象
$ConnectionString = "Server=$serverName;Database=$databaseName;User Id=$userName;Password=$password;"
$SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $ConnectionString
$SqlConnection.Open()

# 获取表格的所有列名
$GetColumnsCommand = $SqlConnection.CreateCommand()
$GetColumnsCommand.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '$schemaName' AND TABLE_NAME = '$tableName'"
$ColumnReader = $GetColumnsCommand.ExecuteReader()
$Columns = @()
while ($ColumnReader.Read()) {
    $Columns += $ColumnReader["COLUMN_NAME"]
}; $SqlConnection.Close()

# 创建一个将所有列名包裹在双引号中的 SQL 查询
$QuotedFields = "SELECT "+"'`"" +($Columns -join "`"','`"")+ "`"'" + ", `"中文測試`""
$QuotedColumns = $Columns -replace ('^(.+)$', '''"'' + REPLACE(CONVERT(NVARCHAR(MAX), $1), ''"'', ''""'') + ''"'' AS $1')
$Query  = "SET NOCOUNT ON`r`n"
$Query += "$QuotedFields`r`n"
$Query += "SELECT`r`n    $($QuotedColumns -join ",`r`n    ")`r`nFROM $Table;"

# 組裝Query
$sqlFile = "sql\tmp.sql"
$Query | Set-Content -Encoding utf8 $sqlFile

# 下載
sqlcmd -S $serverName -d $databaseName -U $userName -P $password -i $sqlFile -o "csv\$tableName.csv" -b -s ',' -W -h -1 -f 65001
