# 拆分表名
function Split-SqlTableName {
    param (
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [string]$TableName
    )

    # 拆分表名並提取資料庫名，模式名和表名
    $splitTable = $TableName.Split('.')
    $databaseName = $null
    $schemaName = $null
    $tableName = $null

    switch ($splitTable.Length) {
        1 {
            $tableName = $splitTable[0] -Replace("^\[|\]$")
        }
        2 {
            $schemaName = $splitTable[0] -Replace("^\[|\]$")
            $tableName = $splitTable[1] -Replace("^\[|\]$")
        }
        3 {
            $databaseName = $splitTable[0] -Replace("^\[|\]$")
            $schemaName = $splitTable[1] -Replace("^\[|\]$")
            $tableName = $splitTable[2] -Replace("^\[|\]$")
        }
        default {
            return $null
        }
    }

    # 拼接完整表名
    $fullTableName = [string]::Empty
    if ($databaseName) {
        $fullTableName += "[$databaseName]."
    }
    if ($schemaName) {
        $fullTableName += "[$schemaName]."
    }
    if ($tableName) {
        $fullTableName += "[$tableName]"
    }

    # 返回包含完整表名的 PSCustomObject
    return [PSCustomObject]@{
        DatabaseName = $databaseName
        SchemaName = $schemaName
        TableName = $tableName
        FullTableName = $fullTableName
    }
}



# 下載MSSQL表的CSV檔案
function Export-MssqlCsv {
    param (
        [string] $ServerName,
        [string] $UserName,
        [string] $Passwd,
        [string] $Table,
        [string] $OutputCSVPath,
        [System.Text.Encoding] $Encoding = (New-Object System.Text.UTF8Encoding $False)
    )
    # 獲取 [資料庫名, 模式名, 表名]
    $tableInfo = Split-SqlTableName $Table
    $DatabaseName = $tableInfo.DatabaseName
    $SchemaName = $tableInfo.SchemaName
    $TableName = $tableInfo.TableName

    # 建立連接到資料庫的 SqlConnection 物件
    $connectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$UserName;Password=$Passwd;"
    $sqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $connectionString
    $sqlConnection.Open()

    # 獲取表格的所有列名
    $getColumnsCommand = $sqlConnection.CreateCommand()
    $getColumnsCommand.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '$SchemaName' AND TABLE_NAME = '$TableName'"
    $columnReader = $getColumnsCommand.ExecuteReader()
    $columns = @()
    while ($columnReader.Read()) {
        $columns += $columnReader["COLUMN_NAME"]
    }
    $sqlConnection.Close()

    # 創建一個將所有列名包裹在雙引號中的 SQL 查詢
    $quotedFields = "SELECT "+"'`"" +($Columns -join "`"','`"")+ "`"'"
    $quotedColumns = $columns -replace ('^(.+)$', '''"'' + REPLACE(CONVERT(NVARCHAR(MAX), $1), ''"'', ''""'') + ''"'' AS $1')
    $query  = "SET NOCOUNT ON`r`n"
    $query += "$quotedFields`r`n"
    $query += "SELECT`r`n    $($quotedColumns -join ",`r`n    ")`r`nFROM $Table;"

    # 輸出 QueryString 檔案
    $sqlFile = "sql\tmp.sql"
    $query | Set-Content -Encoding utf8 $sqlFile

    # 下載
    sqlcmd -S $ServerName -U $UserName -P $Passwd -i $sqlFile -o $OutputCSVPath -b -s ',' -W -h -1 -f ($Encoding.CodePage)
    
    if ($LASTEXITCODE -ne 0) {
        $Content = (Get-Content -Path $OutputCSVPath) -join ", "
        Write-Error $Content
    } else {
        Write-Host "成功: SQL 執行成功完成, 已下載完指定表格"
    }
} # Export-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "CHG.CHG.Employees" -OutputCSVPath "csv\Employees.csv"
