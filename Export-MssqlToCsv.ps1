# 拆分表名
function Split-SqlTableName {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string]$TableName
    )
    
    process {
        # 拆分表名並提取資料庫名，模式名和表名
        $parts = "$TableName".Split('.') -replace '^\[|\]$'
        
        $result = switch ($parts.Count) {
            1 { @($null, $null, $parts[0]); break }
            2 { @($null, $parts[0], $parts[1]); break }
            3 { @($parts[0], $parts[1], $parts[2]); break }
            default { 
                Write-Error "Invalid table name format: $TableName" -ErrorAction $ErrorActionPreference
                return
            }
        }

        # 返回包含完整表名的 PSCustomObject
        [PSCustomObject]@{
            DatabaseName = $result[0]
            SchemaName = $result[1]
            TableName = $result[2]
            FullTableName = '[{0}]' -f ($result.Where({$_}) -join '].[')
        }
    }
} # "[CHG].[CHG].[TEST]", "CHG.CHG.TEST2", "CHG.TEST3", "TEST4" | Split-SqlTableName

# 從SQL查詢獲取結果並輸出到管道 (極簡流式處理版本)
function Get-SqlQueryResult {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory)]
        [string]$ConnectionString,
        
        [Parameter(Position = 1, Mandatory)]
        [string]$Query,
        
        [Parameter()]
        [switch]$Raw
    )
    
    try {
        # 建立連接
        $conn = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $conn.Open()
        
        # 建立命令
        $cmd = New-Object System.Data.SqlClient.SqlCommand($Query, $conn)
        
        # 建立資料讀取器，並使用 SequentialAccess 模式提高效能
        $reader = $cmd.ExecuteReader([System.Data.CommandBehavior]::SequentialAccess)
        if (-not $reader.HasRows) { Write-Verbose "Query did not return any data"; return }
        
        # 獲取欄位名稱
        $fieldCount = $reader.FieldCount
        $columnNames = 0..($fieldCount-1) | ForEach-Object { $reader.GetName($_) }
        if ($Raw) { ,$columnNames }
        
        # 預先分配數組
        $values = New-Object object[] $fieldCount
        
        # 讀取資料流輸出
        while ($reader.Read()) {
            # 一次獲取整行數據
            [void]$reader.GetValues($values)
            
            # 高性能模式：直接輸出數組
            if ($Raw) { ,$values.Clone(); continue }
            
            # 標準模式：輸出屬性雜湊表物件
            $properties = [ordered]@{}
            for ($i = 0; $i -lt $fieldCount; $i++) {
                $properties[$columnNames[$i]] = if ($values[$i] -eq [DBNull]::Value) { $null } else { $values[$i] }
            }; [PSCustomObject]$properties
        }
    }
    catch {
        Write-Error -ErrorRecord $_
    }
    finally {
        # 釋放資源
        if ($reader) { $reader.Dispose() }
        if ($cmd) { $cmd.Dispose() }
        if ($conn) { $conn.Dispose() }
    }
}

# 轉換資料為CSV字串
function ConvertTo-CsvString {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [object[]]$RawData,
        
        [Parameter()]
        [string]$NullValue = 'NULL',
        
        [Parameter()]
        [string]$DateTimeFormat = 'yyyy-MM-dd HH:mm:ss'
    )
    
    begin {
        # CSV 格式化規則
        $csvRules = @{
            # 使用兩個正則表達式組合來實現：
            # 1. 匹配需要引號的基本條件（逗號、換行、引號）
            # 2. 匹配前後空白，但排除全形空白
            NeedsQuotes = '[,\r\n"]|^[\t\n\v\f\r ]|[\t\n\v\f\r ]$'
            QuoteChar = '"'           # 引號字元
            EscapeChar = '""'         # 引號轉義方式
            Delimiter = ','           # CSV 分隔符
        }
    }
    
    process {
        # 使用數組方法處理數據，提高性能
        $values = New-Object string[] $RawData.Length
        
        for ($i = 0; $i -lt $RawData.Length; $i++) {
            $item = $RawData[$i]
            
            # 處理 NULL 值
            if ($null -eq $item -or $item -is [System.DBNull]) {
                $values[$i] = $NullValue
            }
            # 處理日期時間
            elseif ($item -is [DateTime]) {
                $values[$i] = $item.ToString($DateTimeFormat)
            }
            # 處理字串轉換與引號處理
            else {
                $strValue = $item.ToString()
                if ($strValue -match $csvRules.NeedsQuotes) {
                    $values[$i] = $csvRules.QuoteChar + ($strValue -replace $csvRules.QuoteChar, $csvRules.EscapeChar) + $csvRules.QuoteChar
                } else {
                    $values[$i] = $strValue
                }
            }
        }
        
        # 使用高效的 .NET Join 方法連接所有值
        [string]::Join($csvRules.Delimiter, $values)
    }
}

# 測試用指令
function Test-SqlQueryResult {
    [CmdletBinding()]
    param()
    
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
    $data | ConvertTo-CsvString -DateTimeFormat 'yyyy-MM-dd HH:mm:ss.fff'| Set-Content "tmp\CHG.CHG.Table02.csv"
} # Test-SqlQueryResult

# 匯出MSSQL表的CSV檔案
function Export-MssqlToCsv {
    [CmdletBinding(DefaultParameterSetName = "")]
    param (
    )
}
