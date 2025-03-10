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
            
            # 添加標記屬性，表示這是由轉換器創建的
            $connection | Add-Member -NotePropertyName CreatedByTransformer -NotePropertyValue $true -Force
            
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
        
        # 檢查連接是否由轉換器創建，如果是則關閉並釋放
        if ($Connection.PSObject.Properties.Name -contains 'CreatedByTransformer' -and $Connection.CreatedByTransformer) {
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
    
    # 建立連接字串
    # $info = @{
    #     Server   = "UX533-PC"
    #     Database = "CHG"
    #     UserId   = "chg"
    #     Password = "1230"
    # }
    # $conn = "Server={0};Database={1};User Id={2};Password={3}" -f `
    #     $info.Server, $info.Database, $info.UserId, $info.Password
    
    # $conn = [Data.SqlClient.SqlConnection](
    #     ([Data.SqlClient.SqlConnectionStringBuilder]@{
    #         DataSource     = 'UX533-PC'
    #         InitialCatalog = 'CHG'
    #         UserID         = 'chg'
    #         Password       = '1230'
    #     }).ConnectionString
    # )
    
    # $conn = [Data.SqlClient.SqlConnection](
    #     ([Data.SqlClient.SqlConnectionStringBuilder]@{
    #         DataSource     = 'UX533-PC'
    #         InitialCatalog = 'CHG'
    #         UserID         = 'chg'
    #         Password       = '1230'
    #     }).ConnectionString
    # )
    
    $conn = @{
        DataSource     = 'UX533-PC'
        InitialCatalog = 'CHG'
        UserID         = 'chg'
        Password       = '1230'
    }
    
    # $conn = '
    #     Data Source=UX533-PC;
    #     Initial Catalog=CHG;
    #     User ID=chg;
    #     Password=1230
    # '

    # 查詢字串
    $query = "SELECT * FROM [CHG].[CHG].[Table02]"
    
    # 執行查詢
    $data = Get-SqlQueryResult -Connection $conn -Query $query -Verbose -Raw
    
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
    [CmdletBinding(DefaultParameterSetName = "ConnectionHash")]
    param (
        [Parameter(Position = 0, Mandatory, ParameterSetName = "ConnectionHash")]
        [hashtable]$ConnectionInfo,
        
        [Parameter(Position = 0, Mandatory, ParameterSetName = "ConnectionString")]
        [string]$ConnectionString,
        
        [Parameter(Position = 1, Mandatory, ValueFromPipeline)]
        [string[]]$TableName,
        
        [Parameter(Position = 2)]
        [string]$OutputPath,
        
        [Parameter()]
        [string]$NullValue = "NULL",
        
        [Parameter()]
        [string]$DateTimeFormat = "yyyy-MM-dd HH:mm:ss",
        
        [Parameter()]
        [switch]$NoHeaders,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        # 處理連接字串
        if ($PSCmdlet.ParameterSetName -eq "ConnectionHash") {
            # 驗證必要的連接資訊
            if (-not $ConnectionInfo.ContainsKey('Server') -or -not $ConnectionInfo.ContainsKey('Database')) {
                throw "連接資訊必須包含 'Server' 和 'Database' 鍵值"
            }
            
            # 建立連接字串
            $connBuilder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
            $connBuilder['Data Source'] = $ConnectionInfo['Server']
            $connBuilder['Initial Catalog'] = $ConnectionInfo['Database']
            
            # 設定認證 (SQL 或 Windows 整合認證)
            if ($ConnectionInfo.ContainsKey('UserId') -and $ConnectionInfo.ContainsKey('Password')) {
                $connBuilder['User ID'] = $ConnectionInfo['UserId']
                $connBuilder['Password'] = $ConnectionInfo['Password']
            } else {
                $connBuilder['Integrated Security'] = $true
            }
            
            # 設定其他屬性
            $ConnectionInfo.GetEnumerator() | Where-Object { 
                $_.Key -notin @('Server', 'Database', 'UserId', 'Password') 
            } | ForEach-Object {
                try { $connBuilder[$_.Key] = $_.Value } catch { Write-Verbose "忽略屬性 '$($_.Key)'" }
            }
            
            $ConnectionString = $connBuilder.ToString()
        }
        
        Write-Verbose "連接字串: $ConnectionString"
        
        # 如果輸出路徑為空，則使用當前目錄
        if ([string]::IsNullOrWhiteSpace($OutputPath)) {
            $OutputPath = (Get-Location).Path
            Write-Verbose "未指定輸出路徑，使用當前目錄: $OutputPath"
        }
        
        # 確保輸出路徑存在
        if (-not (Test-Path -Path $OutputPath -PathType Container)) {
            try {
                New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
                Write-Verbose "已創建輸出目錄: $OutputPath"
            } catch {
                throw "無法創建輸出目錄 '$OutputPath': $_"
            }
        }
    }
    
    process {
        foreach ($table in $TableName) {
            try {
                # 解析表名
                $parsedTable = $table | Split-SqlTableName
                Write-Verbose "處理表: $($parsedTable.FullTableName)"
                
                # 構建輸出文件名
                $fileName = if ($parsedTable.DatabaseName) {
                    "$($parsedTable.DatabaseName).$($parsedTable.SchemaName).$($parsedTable.TableName).csv"
                } elseif ($parsedTable.SchemaName) {
                    "$($parsedTable.SchemaName).$($parsedTable.TableName).csv"
                } else {
                    "$($parsedTable.TableName).csv"
                }
                $outputFile = Join-Path -Path $OutputPath -ChildPath $fileName
                
                # 檢查文件是否存在
                if ((Test-Path -Path $outputFile) -and -not $Force) {
                    Write-Warning "文件 '$outputFile' 已存在，使用 -Force 參數覆蓋"
                    continue
                }
                
                # 構建查詢
                $query = "SELECT * FROM $($parsedTable.FullTableName)"
                Write-Verbose "執行查詢: $query"
                
                # 執行查詢並獲取數據
                $data = Get-SqlQueryResult -ConnectionString $ConnectionString -Query $query -Raw
                
                if (-not $data) {
                    Write-Warning "表 $($parsedTable.FullTableName) 未返回任何數據"
                    continue
                }
                
                # 處理標題行
                if (-not $NoHeaders -and $data.Count -gt 0) {
                    # 包含標題行
                    $csvLines = @($data[0] | ConvertTo-CsvString -NullValue $NullValue -DateTimeFormat $DateTimeFormat)
                    $startIndex = 1
                } else {
                    # 不包含標題行
                    $csvLines = @()
                    $startIndex = 0
                }
                
                # 處理數據行
                for ($i = $startIndex; $i -lt $data.Count; $i++) {
                    $csvLines += $data[$i] | ConvertTo-CsvString -NullValue $NullValue -DateTimeFormat $DateTimeFormat
                }
                
                # 輸出到文件
                Set-Content -Path $outputFile -Value $csvLines -Encoding UTF8
                
                # 計算行數（不包括標題行）
                $rowCount = $data.Count - $startIndex
                Write-Verbose "已導出 $rowCount 行數據到 '$outputFile'"
                
                # 返回結果對象
                [PSCustomObject]@{
                    TableName = $parsedTable.FullTableName
                    OutputFile = $outputFile
                    RowCount = $rowCount
                }
            }
            catch {
                Write-Error "處理表 '$table' 時出錯: $_"
            }
        }
    }
}

# 使用範例
# $connInfo = @{
#     Server = "UX533-PC"
#     Database = "CHG"
#     UserId = "chg"
#     Password = "1230"
# }

# 匯出單一表格
# Export-SqlServerTableToCsv -ConnectionInfo $connInfo -TableName "[CHG].[CHG].[Table02]" -Force

# 匯出多個表格
# "[CHG].[CHG].[Table02]", "[CHG].[CHG].[Table01]" | 
#     Export-SqlServerTableToCsv -ConnectionInfo $connInfo -OutputPath ".\tmp" -NoHeaders -Force
