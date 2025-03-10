# 自訂型別轉換類別
class SqlConnectionTransformationAttribute : System.Management.Automation.ArgumentTransformationAttribute {
    [object] Transform([System.Management.Automation.EngineIntrinsics] $engineIntrinsics, [object] $inputData) {
        # 如果已經是 SqlConnection，直接返回
        if ($inputData -is [System.Data.SqlClient.SqlConnection]) {
            return $inputData
        }
        
        # 檢查輸入型態是否為支援的型態
        if (-not ($inputData -is [hashtable] -or 
                  $inputData -is [System.Data.SqlClient.SqlConnectionStringBuilder] -or 
                  $inputData -is [string])) {
            throw [System.ArgumentException]::new("Cannot convert value of type '$($inputData.GetType().FullName)' to type 'System.Data.SqlClient.SqlConnection'. The argument type must be SqlConnection, Hashtable, SqlConnectionStringBuilder, or String.")
        }
        
        # 嘗試轉換輸入型態
        try {
            # 處理雜湊表
            if ($inputData -is [hashtable]) {
                $inputData = [System.Data.SqlClient.SqlConnectionStringBuilder]$inputData
            }
            # 處理 SqlConnectionStringBuilder
            if ($inputData -is [System.Data.SqlClient.SqlConnectionStringBuilder]) {
                $inputData = $inputData.ConnectionString
            }
            # 處理字串和其他類型
            $connection = [System.Data.SqlClient.SqlConnection]::new($inputData.ToString())
            
            # 獲取調用者的名稱
            $callerName = $engineIntrinsics.SessionState.PSVariable.GetValue('PSCmdlet').MyInvocation.MyCommand.Name
            
            # 添加標記屬性，用於追蹤連接的擁有者
            $connection | Add-Member -NotePropertyName CallerName -NotePropertyValue $callerName -Force
            
            return $connection
        } catch { throw }
    }
}

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
        [SqlConnectionTransformation()]
        [Data.SqlClient.SqlConnection]$Connection,
        
        [Parameter(Position = 1, Mandatory)]
        [string]$Query,
        
        [Parameter()]
        [switch]$Raw
    )
    
    try {
        # 開啟連接
        if ($Connection.State -ne [System.Data.ConnectionState]::Open) { $Connection.Open() }
        $cmd = New-Object System.Data.SqlClient.SqlCommand($Query, $Connection)
        $reader = $cmd.ExecuteReader([System.Data.CommandBehavior]::SequentialAccess)
        if (-not $reader.HasRows) { Write-Verbose "Query did not return any data"; return }
        
        # 獲取欄位名稱
        $fieldCount = $reader.FieldCount
        $columnNames = 0..($fieldCount-1) | ForEach-Object { $reader.GetName($_) }
        if ($Raw) { ,$columnNames }
        
        # 讀取資料流輸出
        $values = New-Object object[] $fieldCount
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
        Write-Error $_
    }
    finally {
        # 釋放資源
        if ($reader) { $reader.Dispose() }
        if ($cmd) { $cmd.Dispose() }
        
        # 檢查連接是否由當前函式擁有，如果是則關閉並釋放
        $currentFunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
        if ($Connection.PSObject.Properties.Name -contains 'CallerName' -and $Connection.CallerName -eq $currentFunctionName) {
            if ($Connection.State -ne [System.Data.ConnectionState]::Closed) {
                $Connection.Close()
            }; $Connection.Dispose()
        }
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
    
    # 執行查詢
    $data = Get-SqlQueryResult -Connection @{
        DataSource     = 'UX533-PC'
        InitialCatalog = 'CHG'
        UserID         = 'chg'
        Password       = '1230'
    } -Query "SELECT * FROM [CHG].[CHG].[Table02]" -Verbose -Raw
    
    # 輸出到CSV
    $data | ConvertTo-CsvString -DateTimeFormat 'yyyy-MM-dd HH:mm:ss.fff'| Set-Content "tmp\CHG.CHG.Table02.csv" -Encoding utf8BOM
    
    # 計算檔案雜湊值進行比對
    $expectedHash = Get-FileHash -Path "tmp\CHG.CHG.Table02.csv" -Algorithm SHA256
    $actualHash = Get-FileHash -Path "test-csv-output-data\ssms-output.csv" -Algorithm SHA256
    if ($expectedHash.Hash -ne $actualHash.Hash) {
        Write-Host "檔案內容不一致" -ForegroundColor Red
    } else { Write-Host "檔案內容一致" -ForegroundColor Green }
} # Test-SqlQueryResult

# 匯出MSSQL表的CSV檔案
function Export-SqlServerTableToCsv {
    [CmdletBinding(DefaultParameterSetName = "")]
    param (
        # 連接字串
        [Parameter(Position = 0, Mandatory)]
        [SqlConnectionTransformation()]
        [Data.SqlClient.SqlConnection]$Connection,
        # 表格名稱
        [Parameter(Position = 1, Mandatory)]
        [string]$TableName,
        
        # 輸出CSV檔案路徑 (預設$NULL會取當前工作目錄)
        [Parameter(Position = 2)]
        [string]$Path,
        # 處理NULL值
        [Parameter()]
        [string]$NullValue = "NULL",
        # 日期時間格式
        [Parameter()]
        [string]$DateTimeFormat = "yyyy-MM-dd HH:mm:ss",
        # 不輸出CSV檔案標頭
        [Parameter()]
        [switch]$NoHeaders,
        # 強制覆蓋CSV檔案
        [Parameter()]
        [switch]$Force
    )
    
    try {
        # 解析表格名稱
        $parsedTable = $TableName | Split-SqlTableName
        if (-not $parsedTable) { return }

        # 處理路徑
        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if (-not [IO.Path]::HasExtension($Path)) {
            $fileName = ($parsedTable.FullTableName -replace '\[|\]' -replace '\.', '_') + '.csv'
            $Path = Join-Path $Path $fileName
        }
        
        # 檢查檔案是否已存在
        if ((Test-Path $Path) -and (-not $Force)) {
            Write-Error "檔案 '$Path' 已存在。使用 -Force 參數來覆蓋檔案。"
            return
        }
        
        # 構建查詢
        $query = "SELECT * FROM $($parsedTable.FullTableName)"
        
        # 創建一個計數器
        $rowCount = 0
        
        # 執行查詢並將結果儲存到CSV檔案，同時計算行數
        Get-SqlQueryResult -Connection $Connection -Query $query -Raw | 
            ForEach-Object { $rowCount++; $_ } |
            ConvertTo-CsvString -NullValue $NullValue -DateTimeFormat $DateTimeFormat | 
            Set-Content -Path $Path -Encoding utf8BOM
        
        # 返回結果對象
        [PSCustomObject]@{
            TableName  = $parsedTable.FullTableName
            OutputFile = $Path
            RowCount   = $rowCount
        }
    }
    finally {
        # 檢查連接是否由當前函式擁有，如果是則關閉並釋放
        $currentFunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
        if ($Connection.PSObject.Properties.Name -contains 'CallerName' -and $Connection.CallerName -eq $currentFunctionName) {
            if ($Connection.State -ne [System.Data.ConnectionState]::Closed) {
                $Connection.Close()
            }
            $Connection.Dispose()
        }
    }
}

# 使用範例
# $dateTimeFormat = 'yyyy-MM-dd HH:mm:ss.fff'
# $cnnInfo = @{
#     DataSource     = 'UX533-PC'
#     InitialCatalog = 'CHG'
#     UserID         = 'chg'
#     Password       = '1230'
# }
# $tableName = "[CHG].[CHG].[Table02]"

# 完整輸出檔名
# Export-SqlServerTableToCsv $cnnInfo $tableName "tmp\CHG.CHG.Table02.csv" -DateTimeFormat $dateTimeFormat -Force

# 輸入目錄(自動取表名)
# Export-SqlServerTableToCsv $cnnInfo $tableName "tmp" -DateTimeFormat $dateTimeFormat -Force

# 輸出到當前目錄(自動取表名)
# Export-SqlServerTableToCsv $cnnInfo $tableName -DateTimeFormat $dateTimeFormat -Force
