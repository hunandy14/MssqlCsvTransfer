# 移除 CSV 數據中的雙引號
function Remove-CsvQuotes {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $InputPath,
        [Parameter(ParameterSetName = "")]
        [string] $OutputPath,
        [Parameter(ParameterSetName = "")]
        [text.encoding] $Encoding = (New-Object System.Text.UTF8Encoding $False),
        [switch] $RemoveHeader
    )
    [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
    
    # 設定分隔符號
    [string] $Delimiter = ','

    # 確認輸入檔案存在
    if (!(Test-Path -PathType:Leaf $InputPath)) { Write-Error "The [`$InputPath:: `"$InputPath`"] does not exist."; return }
    $InputPath = [IO.Path]::GetFullPath($InputPath)
    # 創建輸出檔案
    if ($OutputPath) {
        $OutputPath = [IO.Path]::GetFullPath($OutputPath)
        if (!(Test-Path $OutputPath)) { New-Item $OutputPath -Force -EA:Stop | Out-Null }
    } else { $OutputPath = New-TemporaryFile }
    
    # 流處理檔案
    $headerProcessed = $false
    $writer = New-Object System.IO.StreamWriter -ArgumentList $OutputPath, $false, $Encoding
    $reader = New-Object System.IO.StreamReader -ArgumentList $InputPath, $Encoding
    if ($reader) {
        while (!$reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ($RemoveHeader -and !$headerProcessed) {
                $headerProcessed = $true
                continue
            }
            $fields = $line.Split($Delimiter)
            $newLine = ""
            for ($i = 0; $i -lt $fields.Length; $i++) {
                $cleanField = $fields[$i] -replace('^"|"$') # -replace('"', '""')
                if ($i -gt 0) { $newLine += $Delimiter }
                $newLine += $cleanField
            }; $writer.WriteLine($newLine)
        }; $reader.Close()
    }; $writer.Close()
    
    return $OutputPath
} # Remove-CsvQuotes -InputPath 'csv\Data.csv' -OutputPath 'data\Data.data' -RemoveHeader
# $tmp = Remove-CsvQuotes -InputPath 'csv\Data.csv' -RemoveHeader; if ($tmp) { $tmp; Remove-Item "$($tmp -replace '.tmp$').tmp" }



# 上傳 CSV 到 MSSQL資料庫
function Import-MssqlCsv {
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName = "", Mandatory)]
        [string] $ServerName,
        [Parameter(ParameterSetName = "", Mandatory)]
        [string] $UserName,
        [Parameter(ParameterSetName = "", Mandatory)]
        [string] $Passwd,
        [Parameter(ParameterSetName = "", Mandatory)]
        [string] $Table,
        [Parameter(ParameterSetName = "", Mandatory)]
        [string] $CsvPath,
        [Parameter(ParameterSetName = "")]
        [text.encoding] $Encoding = (New-Object System.Text.UTF8Encoding $False),
        [switch] $NonHeaderFile
    )
    
    begin {
        $tmp = New-TemporaryFile
        $dataPath = $tmp.FullName
        Remove-CsvQuotes -InputPath $CsvPath -OutputPath $dataPath -RemoveHeader:(!$NonHeaderFile) -Encoding:$Encoding | Out-Null
        $CsvPath = $dataPath
        [string] $Terminator = ','
        [string] $RowTerminator = "`r`n"
    }
    
    process {
        $Output = & bcp $Table in $CsvPath -C ($Encoding).CodePage -c -t $Terminator -r $RowTerminator -S $ServerName -U $UserName -P $Passwd
        $HasError = $false
        $RowsCopied = 0
        $OutputString = $Output -join "`r`n"
        if ($outputString -match "(\d+) rows copied\.") {
            $RowsCopied = [int]$matches[1]
            if ($RowsCopied -eq 0) { $HasError = $true }
        } else { $HasError = $true }
    }
    
    end {
        # 刪除暫存檔案
        if ($tmp) {
            $tmpPath = $tmp.FullName -replace '.tmp$'
            Remove-Item "$tmpPath.tmp"
        }
        # 回傳物件
        return [pscustomobject]@{
            IsSuccessful = !$HasError
            RowsCopied   = $RowsCopied
            Message      = $Output -match ".+" -notmatch "Starting copy..."
        }
    }
} # Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -CsvPath "csv\Data.csv"
# Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -CsvPath "data\Data.data" -NonHeaderFile
