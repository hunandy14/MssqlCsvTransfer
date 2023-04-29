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
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $ServerName,
        [Parameter(Position = 1, ParameterSetName = "", Mandatory)]
        [string] $UserName,
        [Parameter(Position = 2, ParameterSetName = "", Mandatory)]
        [string] $Passwd,
        [Parameter(Position = 3, ParameterSetName = "Table", Mandatory)]
        [string] $Table,
        [Parameter(ParameterSetName = "Table")]
        [string] $HeaderString,
        [Parameter(Position = 3, ParameterSetName = "SQL", Mandatory)]
        [string] $SQLPath,
        [Parameter(ParameterSetName = "")]
        [string] $Path,
        [Text.Encoding] $Encoding = (New-Object System.Text.UTF8Encoding $False),
        [Parameter(ParameterSetName = "")]
        [switch] $OutNull,
        [switch] $OutToTemp,
        [switch] $OpenOutDir
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

    # 路徑處理
    [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
    if (!$Path) {
        if ($OutToTemp) { # 輸出到臨時檔案
            $Path = $env:TEMP + "\Export-MssqlCsv\$TableName.csv"
        } else { # 輸出到當前資料夾
            $Path = [IO.Path]::GetFullPath("$TableName.csv")
        }
    } else { $Path = [IO.Path]::GetFullPath($Path) }
    
    # 創建空檔案
    if (!(Test-Path $Path)) { New-Item $Path -ItemType:File -Force -EA:Stop | Out-Null }
    
    if (!$SQLPath) {
        # 建立連接到資料庫的 SqlConnection 物件
        $connectionString = "Server=$ServerName;Database=$DatabaseName;User Id=$UserName;Password=$Passwd;"
        $sqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $connectionString
        try {
            $sqlConnection.Open()
        } catch { Write-Error "Unable to open SQL connection: $_" -EA:Stop }
    
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
        $quotedColumns = $columns -replace ('^(.+)$', '''"'' + REPLACE($1, ''"'', ''""'') + ''"'' AS $1')
        $query  = "SET NOCOUNT ON`r`n"
        $query += "$quotedFields`r`n"
        $query += "SELECT`r`n    $($quotedColumns -join ",`r`n    ")`r`nFROM $Table;"
    
        # 輸出 QueryString 檔案
        # $tmp = New-TemporaryFile
        # $sqlFile = $tmp.FullName
        # $sqlFile = $env:TEMP + "\Export-MssqlCsv\Download.sql"
        $sqlFile = $env:TEMP + "Download.sql"
        if (!(Test-Path $sqlFile)) { New-Item $sqlFile -ItemType:File -Force | Out-Null }
        $sqlFile
        $query | Set-Content -Encoding utf8 $sqlFile
    } else {
        $sqlFile = $SQLPath
    }
       
    # 下載
    sqlcmd -S $ServerName -U $UserName -P $Passwd -i $sqlFile -o $Path -b -s ',' -W -h -1 -f ($Encoding.CodePage)

    # 刪除暫存SQL檔案
    if ($tmp) {
        $tmpPath = $tmp.FullName -replace '.tmp$'
        Remove-Item "$tmpPath.tmp"
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
            Write-Host "Success:: SQL execution completed, `"$FullTableName`" has been downloaded to"
            Write-Host "  $Path" -ForegroundColor Yellow
        }
        if ($OpenOutDir) {
            explorer.exe (Split-Path $Path -Parent)
        }
        return $Path
    }
} # Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "CHG.CHG.Employees" | Out-Null
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "CHG.CHG.Employees" -HeaderString '"EmployeeID","FirstName","LastName","BirthDate"' -OutToTemp | Out-Null
# Export-MssqlCsv "192.168.3.123,1433" "kaede" "1230"  -SQLPath "sql\V03.sql" | Out-Null
