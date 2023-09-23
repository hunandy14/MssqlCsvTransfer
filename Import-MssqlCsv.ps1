# 移除 CSV 數據中的雙引號
function Remove-CsvQuotes {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $Path,            # 輸入的帶有雙引號的CSV數據
        [Parameter(ParameterSetName = "")]
        [string] $Output,
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM,
        [Parameter(ParameterSetName = "")]
        [switch] $RemoveHeader,    # 移除CSV的字段
        [Parameter(ParameterSetName = "")]
        [string] $ReplaceDelimiter # 替換CSV的分隔符號
    )
    
    # 設定預設分隔符號
    [string] $Delimiter = ','
    
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

    # 確認輸入檔案存在
    if (!(Test-Path -PathType:Leaf $Path)) { Write-Error "The [`$Path:: `"$Path`"] does not exist."; return }
    $Path = [IO.Path]::GetFullPath([IO.Path]::Combine((Get-Location -PSProvider FileSystem).ProviderPath, $Path))
    # 創建輸出檔案
    if ($Output) {
        $Output = [IO.Path]::GetFullPath($Output)
        if (!(Test-Path $Output)) { New-Item $Output -Force -EA:Stop | Out-Null }
    } else { $Output = New-TemporaryFile }
    
    # 流處理檔案
    $headerProcessed = $false
    $writer = New-Object System.IO.StreamWriter -ArgumentList $Output, $false, $Encoding
    $reader = New-Object System.IO.StreamReader -ArgumentList $Path, $Encoding
    if ($reader) {
        while (!$reader.EndOfStream) {
            $line = $reader.ReadLine()
            if ($RemoveHeader -and !$headerProcessed) {
                $headerProcessed = $true
                continue
            }
            # 更換分隔符號
            if ($ReplaceDelimiter) {
                $line = $line -replace "($Delimiter)(?=(?:[^""]*""[^""]*"")*[^""]*$)", $ReplaceDelimiter
                $delim = $ReplaceDelimiter
            } else { $delim = $Delimiter }
            # 消除雙引號
            $line = $line -replace "(?<=^|\s*$delim\s*)""\s*|\s*""(?=\s*$delim|$)" -replace '""', '"'
            # 寫入檔案
            $writer.WriteLine($line)
        }; $reader.Close()
    }; $writer.Close()
    
    return $Output
} # Remove-CsvQuotes 'csv\Data.csv' -Output 'data\Data.data' -RemoveHeader
# $tmp = Remove-CsvQuotes 'csv\Data.csv' -RemoveHeader; if ($tmp) { $tmp; Remove-Item "$($tmp -replace '.tmp$').tmp" }
# Remove-CsvQuotes 'csv\Data.csv' -Output 'data\Data.csv' -RemoveHeader -ReplaceDelimiter '¬' -UTF8BOM



# 上傳 CSV 到 MSSQL資料庫
function Import-MssqlCsv {
    [CmdletBinding()]
    param(
        # 主要參數
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $ServerName,
        [Parameter(Position = 1, ParameterSetName = "", Mandatory)]
        [string] $UserName,
        [Parameter(Position = 2, ParameterSetName = "", Mandatory)]
        [string] $Passwd,
        [Parameter(Position = 3, ParameterSetName = "", Mandatory)]
        [string] $Table,
        [Parameter(Position = 4, ParameterSetName = "", Mandatory)]
        [string] $Path,
        # 前置處理CSV檔案 (會重寫第二份檔案)
        [Parameter(ParameterSetName = "")]
        [switch] $Csv_RemoveQuotes,
        [switch] $Csv_RemoveQuotesHeaders,
        [string] $Csv_ReplaceDelimiter,
        [string] $TempPath, # 使用 '.tmp' 會自動刪除
        # 編碼相關
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM,
        # 其他選項
        [Parameter(ParameterSetName = "")]
        [switch] $CleanTable,
        [switch] $ShowCommand,
        [switch] $OutNull
    )
    
    begin {
        # 設定值
        [string] $Delimiter = ','
        [string] $Terminator = '"`r`n"'
        # 檢查分隔符號是不是 ASCII 字符
        if ($Csv_ReplaceDelimiter -and ($Csv_ReplaceDelimiter -match '[^\x00-\x7F]')) {
            Write-Error "The delimiter '$Csv_ReplaceDelimiter' contains non-ASCII characters. Please use ASCII characters only." -ErrorAction Stop
        }
        # 處理編碼
        if (!$__SysEnc__) { $Script:__SysEnc__ = [Text.Encoding]::GetEncoding((powershell -nop "([Text.Encoding]::Default).WebName")) }
        if (!$Encoding) {
            # 預選項編碼
            if ($UTF8) {
                $Encoding = New-Object System.Text.UTF8Encoding $False
            } elseif ($UTF8BOM) {
                $Encoding = New-Object System.Text.UTF8Encoding $True
            } else { # 系統語言
                $Encoding = $__SysEnc__
            }
        }
        # 確認輸入檔案存在
        if (!(Test-Path -PathType:Leaf $Path)) { Write-Error "The [`$Path:: `"$Path`"] does not exist." -EA:Stop }
        [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
        $Path = [IO.Path]::GetFullPath($Path)
        # 消除檔頭與雙引號
        if ($Csv_RemoveQuotesHeaders -or $Csv_RemoveQuotes -or $Csv_ReplaceDelimiter) {
            $Path = $tmp = Remove-CsvQuotes $Path -Output:$TempPath -RemoveHeader:$Csv_RemoveQuotesHeaders -ReplaceDelimiter:$Csv_ReplaceDelimiter -Encoding:$Encoding
            if ($Csv_ReplaceDelimiter) { $Delimiter = $Csv_ReplaceDelimiter }
        }
    }
    
    process {
        # 獲取編碼號
        $EnvCodePage = $__SysEnc__.CodePage
        $CsvCodePage = $Encoding.CodePage
        # 清空既有的表格
        if ($CleanTable) {
            $cmdStr = "sqlcmd -S $ServerName -U $UserName -P $Passwd -f $EnvCodePage -Q 'DELETE FROM $Table'"
            if ($ShowCommand) { Write-Host $cmdStr -ForegroundColor DarkGray }
            $Result = @()
            (Invoke-Expression $cmdStr) | ForEach-Object {
                if (!$OutNull) { Write-Host $_ }
                $Result += $_
            }; if (!$OutNull) { Write-Host "" }
        }
        # 執行命令 bcp 命令上傳
        $cmdStr = "bcp $Table in '$Path' -C $CsvCodePage -c -t '$Delimiter' -r $Terminator -S $ServerName -U $UserName -P $Passwd"
        if ($ShowCommand) { Write-Host $cmdStr -ForegroundColor DarkGray }
        $Result = @()
        (Invoke-Expression $cmdStr) | ForEach-Object {
            if (!$OutNull) { Write-Host $_ }
            $Result += $_
        }; if (!$OutNull) { Write-Host "" }
        # 獲取上傳結果
        $HasError = $false; $RowsCopied = 0
        if (($Result -join "`r`n") -match "(\d+) rows copied\.") {
            $RowsCopied = [int]$matches[1]
            if ($RowsCopied -eq 0) { $HasError = $true }
        } else { $HasError = $true }
    }
    
    end {
        # 刪除暫存檔案
        if ($tmp -and ($tmp  -match "\.tmp$")) {
            Remove-Item "$($tmp -replace '\.tmp$').tmp"
        }
        # 回傳物件
        return @{
            IsSuccessful = !$HasError
            RowsCopied   = $RowsCopied
            Message      = $Result -match ".+" -notmatch "Starting copy..."
        }
    }
}
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand |Out-Null
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand -Csv_RemoveQuotes |Out-Null
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand -Csv_RemoveQuotesHeaders |Out-Null
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand -CleanTable -Csv_RemoveQuotes -Csv_ReplaceDelimiter '¬'  |Out-Null
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand -CleanTable -Csv_RemoveQuotes -Csv_ReplaceDelimiter '`,' |Out-Null
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand -CleanTable -Csv_RemoveQuotes -Csv_ReplaceDelimiter '`,' -TempPath "data\Data.csv" |Out-Null
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8 -ShowCommand -CleanTable -Csv_RemoveQuotes -Csv_ReplaceDelimiter '`,' -TempPath "data\Data.tmp" |Out-Null
