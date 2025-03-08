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
