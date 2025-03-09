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

# 匯出MSSQL表的CSV檔案
function Export-MssqlToCsv {
    [CmdletBinding(DefaultParameterSetName = "")]
    param (
    )
}

# 從SQL查詢獲取結果並輸出到管道 (極簡流式處理版本)
function Get-SqlQueryResult {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$ConnectionString,
        
        [Parameter(Mandatory)]
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

# 轉換RAW資料為CSV字串
function ConvertTo-CsvString {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$RawData
    )
    
    begin {
        # 用於判斷是否需要引號包覆的函數
        function NeedsQuoting($value) {
            if ($null -eq $value) { return $false }  # NULL 值不需要引號
            if ($value -eq '') { return $false }     # 空字串不需要引號
            
            $strValue = "$value"
            
            # 檢查是否包含需要引號的字元
            return $strValue -match '[,\r\n"]' -or   # 包含逗號、換行或引號
                   $strValue -match '^\s|\s$'        # 前後有空白
        }
        
        # 處理值的函數，根據規則決定是否加引號並處理特殊情況
        function FormatValue($value) {
            # 處理 NULL 值 - 在 SQL 結果中，DBNull.Value 會被轉換為 $null
            if ($null -eq $value -or $value -is [System.DBNull]) { 
                return 'NULL' 
            }
            
            # 處理空字串
            if ($value -eq '') { return '' }
            
            # 處理日期時間格式
            if ($value -is [DateTime]) {
                return $value.ToString("yyyy-MM-dd HH:mm:ss.fff")
            }
            
            # 轉換為字串，保留原始字元（包括控制字元）
            $strValue = "$value"
            
            # 檢查是否需要引號
            if (NeedsQuoting $strValue) {
                # 處理引號：將字串中的每個引號替換為兩個引號
                $quoted = $strValue -replace '"', '""'
                return """$quoted"""  # 加上外層引號
            } else {
                return $strValue  # 不需要引號的情況
            }
        }
    }
    
    process {
        # 處理每個元素並用逗號連接
        ($RawData | ForEach-Object { FormatValue $_ }) -join ','
    }
    
    end {
        # 結束處理
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
    
    # 查詢
    $query = "SELECT * FROM [CHG].[CHG].[Table02]"
    
    # 執行查詢
    $data = Get-SqlQueryResult -ConnectionString $connectionString -Query $query -Verbose -Raw
    
    # 輸出到CSV
    $data | ConvertTo-CsvString | Set-Content -Path "tmp\CHG.CHG.Table02.csv"
    
} Test-SqlQueryResult
