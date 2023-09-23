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
} # "[CHG].[CHG].[TEST]" | Split-SqlTableName



# 下載MSSQL表的CSV檔案
function Export-MssqlCsv {
    [CmdletBinding(DefaultParameterSetName = "Table")]
    param (
        # 登入資訊
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $ServerName,
        [Parameter(Position = 1, ParameterSetName = "", Mandatory)]
        [string] $UserName,
        [Parameter(Position = 2, ParameterSetName = "", Mandatory)]
        [string] $Passwd,
        # 下載源 (表格/Query語句)
        [Parameter(Position = 3, ParameterSetName = "Table", Mandatory)]
        [string] $Table,
        [Parameter(Position = 3, ParameterSetName = "SQLPath", Mandatory)]
        [string] $SQLPath,       # 使用 [-SQLPath] 來啟用這項
        [Parameter(Position = 3, ParameterSetName = "SQLQuery", Mandatory, ValueFromPipeline)]
        [string] $SQLQuery,      # 使用 [-SQLQuery, 管道] 來啟用這項
        # 附加選項
        [Parameter(ParameterSetName = "")]
        [string] $HeaderString,  # 自定義CSV第一行的Header
        [Parameter(ParameterSetName = "")]
        [string] $TempSqlPath,   # 指定自動生成的語句的儲存位置 (未指定是放在temp且會自動刪除)
        # 輸出位置
        [Parameter(ParameterSetName = "")]
        [string] $Path,
        # 編碼
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM,
        # 其他選項
        [Parameter(ParameterSetName = "")]
        [switch] $OutNull,
        [switch] $OutToTemp,
        [switch] $OpenOutDir,
        [switch] $ShowCommand
    )
    
    # 獲取 [資料庫名, 模式名, 表名]
    if ($Table) {
        $tableInfo = Split-SqlTableName $Table
        $DatabaseName = $tableInfo.DatabaseName
        $SchemaName = $tableInfo.SchemaName
        $TableName = $tableInfo.TableName
        $FullTableName = $tableInfo.FullTableName
    } elseif($SQLPath) {
        $FileName = $SQLPath -replace '^(.*[\\/])([^\\/]+?)(\.[^\\/.]+)?$','$2'
        $TableName = $FileName
        $FullTableName = $FileName
    }
    
    # 處理編碼
    if (!$Encoding) {
        # 預選項編碼
        if ($UTF8) {
            $Encoding = New-Object System.Text.UTF8Encoding $False
        } elseif ($UTF8BOM) {
            $Encoding = New-Object System.Text.UTF8Encoding $True
        } else { # 系統語言
            if (!$__SysEnc__) { $Script:__SysEnc__ = [Text.Encoding]::GetEncoding((powershell -nop "([Text.Encoding]::Default).WebName")) }
            $Encoding = $__SysEnc__
        }
    }

    # 路徑處理
    [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
    if (!$Path) {
        $outFileName = $TableName
        if (!$outFileName) { $outFileName = 'QueryResult' }
        if ($OutToTemp) { # 輸出到臨時檔案
            $Path = $env:TEMP + "\Export-MssqlCsv\$outFileName.csv"
        } else { # 輸出到當前資料夾
            $Path = [IO.Path]::GetFullPath("$outFileName.csv")
        }
    } else { $Path = [IO.Path]::GetFullPath($Path) }
    
    # 若輸入非$SQLPath則自動生成SQL文件
    if ($SQLPath) {
        $sqlFile = $SQLPath
    } else {
        # 生成 Query 語句
        if ($SQLQuery) {
            $query = "SET NOCOUNT ON`r`n" + $SQLQuery
        } else {
            # 建立連接到資料庫的 SqlConnection 物件
            $connectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$UserName;Password=$Passwd;"
            $sqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $connectionString
            try {
                $sqlConnection.Open()
            } catch { Write-Error "Unable to open SQL connection: $_" -EA:Stop }
            
            # 檢查表格或檢視是否存在
            $cmdText = "SELECT COUNT(*) FROM sys.objects WHERE object_id = OBJECT_ID(N'$FullTableName') AND type IN (N'U', N'V')"
            $sqlCommand = New-Object -TypeName System.Data.SqlClient.SqlCommand -ArgumentList $cmdText, $sqlConnection
            $tableExists = [int]$sqlCommand.ExecuteScalar()
            if ($tableExists -eq 0) { Write-Error "Table/view '$FullTableName' does not exist." -EA:Stop }
        
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
            if (!$HeaderString) {
                $quotedFields = "SELECT "+"'`"" +($Columns -join "`"','`"")+ "`"'"
            } else { $quotedFields = "SELECT '$HeaderString'" }
            $quotedColumns = $columns -replace ('^(.+)$', '''"'' + REPLACE(ISNULL($1,''''), ''"'', ''""'') + ''"'' AS $1')
            $query  = "SET NOCOUNT ON`r`n"
            $query += "$quotedFields`r`n"
            $query += "SELECT`r`n    $($quotedColumns -join ",`r`n    ")`r`nFROM $Table;"
        }
    
        # 生成空白檔案
        if ($TempSqlPath) {
            $TempSqlPath = [IO.Path]::GetFullPath($TempSqlPath)
            $sqlFile = $tmp = $TempSqlPath
        } else {
            $tmp = New-TemporaryFile
            $sqlFile = $tmp.FullName
        } if (!(Test-Path $sqlFile)) { New-Item $sqlFile -ItemType:File -Force | Out-Null }
        
        # 輸出 QueryString 到檔案
        [IO.File]::WriteAllText($sqlFile, $query, $Encoding)
    }
    
    # 下載
    if (!(Test-Path $Path)) { New-Item $Path -ItemType:File -Force -EA:Stop | Out-Null }
    $cmdStr = "sqlcmd -S '$ServerName' -U '$UserName' -P '$Passwd' -i '$sqlFile' -o '$Path' -b -s ',' -W -h -1 -f $($Encoding.CodePage) -N -C"
    if ($ShowCommand) { Write-Host $cmdStr -ForegroundColor DarkGray }
    $cmdStr | Invoke-Expression

    # 刪除暫存SQL檔案
    if ($tmp -and ($tmp  -match "\.tmp$")) {
        Remove-Item "$($tmp -replace '\.tmp$').tmp"
    }

    # 執行完畢信息處理
    if ($LASTEXITCODE -ne 0) {
        $Content = (Get-Content -Path $Path) -join ", "
        if (!$OutNull) {
            Write-Error $Content -EA:Stop
        }
        return $null
    } else {
        if (!$OutNull) {
            if ($Table) {
                $name = $FullTableName
            } elseif($SQLPath) {
                $name = $sqlFile
            } elseif ($SQLQuery) {
                $name = $SQLQuery
            }
            Write-Host "Success::" -BackgroundColor DarkGreen -ForegroundColor White -NoNewline
            Write-Host " SQL execution completed, '$name' has been downloaded."
        }
        if ($OpenOutDir) {
            explorer.exe (Split-Path $Path -Parent)
        }
        return $Path
    }
} # Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "CHG.CHG.Employees" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "CHG.CHG.Employees" -HeaderString '"EmployeeID","FirstName","LastName","BirthDate"' -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "CHG.CHG.Employees" -TempSqlPath "sql\Employees.sql" -Path "csv\Employees.csv" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "CHG.CHG.Employees" -TempSqlPath "sql\Employees.tmp" -Path "csv\Employees.csv" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLPath "sql\Employees.sql" -Path "csv\Employees.csv" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLPath "sql\V03.sql" -UTF8 -Path "csv\V03.csv"
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLQuery "Select * From CHG.CHG.Employees" -Path "csv\Employees2.csv" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLQuery "Select * From CHG.CHG.Employees Where FirstName = N'あいうえおㄅㄆㄇㄈ'" -TempSqlPath "sql\Employees2.sql" -Path "csv\Employees2.csv" -UTF8
# "Select * From CHG.CHG.Employees" | Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Path "csv\Employees2.csv" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -SQLPath "sql\encrypt.sql" -Path "csv\encrypt.txt" -UTF8
# Export-MssqlCsv "192.168.3.123,1433" "sa" "12301230" -SQLPath "sql\encrypt.sql" -Path "csv\encrypt.txt" -UTF8
# "SET NOCOUNT ON; SELECT DISTINCT (encrypt_option) FROM sys.dm_exec_connections;" | Export-MssqlCsv "192.168.3.123,1433" "sa" "12301230" -UTF8 -ShowCommand
