# 移除 CSV 數據中的雙引號
function Remove-CsvQuotes {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $Path,
        [Parameter(ParameterSetName = "")]
        [string] $Output,
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM,
        [Parameter(ParameterSetName = "")]
        [switch] $RemoveHeader
    )
    
    # 設定分隔符號
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
            $line = $line -replace "(?<=^|\s*$Delimiter\s*)""\s*|\s*""(?=\s*$Delimiter|$)" -replace '""', '"'
            $writer.WriteLine($line)
        }; $reader.Close()
    }; $writer.Close()
    
    return $Output
} # Remove-CsvQuotes 'csv\Data.csv' -Output 'data\Data.data' -RemoveHeader
# $tmp = Remove-CsvQuotes 'csv\Data.csv' -RemoveHeader; if ($tmp) { $tmp; Remove-Item "$($tmp -replace '.tmp$').tmp" }



# 上傳 CSV 到 MSSQL資料庫
function Import-MssqlCsv {
    [CmdletBinding()]
    param(
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
        [switch] $NoHeaders,
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM
    )
    
    begin {
        # 設定值
        [string] $Terminator = ','
        [string] $RowTerminator = "`r`n"
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
        if (!(Test-Path -PathType:Leaf $Path)) { Write-Error "The [`$Path:: `"$Path`"] does not exist." -EA:Stop }
        [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
        $Path = [IO.Path]::GetFullPath($Path)
        # 消除檔頭與雙引號
        $Path = $tmp = Remove-CsvQuotes $Path -RemoveHeader:(!$NoHeaders) -Encoding:$Encoding
    }
    
    process {
        $Result = & bcp $Table in $Path -C ($Encoding).CodePage -c -t $Terminator -r $RowTerminator -S $ServerName -U $UserName -P $Passwd
        $HasError = $false
        $RowsCopied = 0
        if (($Result -join "`r`n") -match "(\d+) rows copied\.") {
            $RowsCopied = [int]$matches[1]
            if ($RowsCopied -eq 0) { $HasError = $true }
        } else { $HasError = $true }
    }
    
    end {
        # 刪除暫存檔案
        if ($tmp) {
            Remove-Item "$($tmp -replace '.tmp$').tmp"
        }
        # 回傳物件
        return @{
            IsSuccessful = !$HasError
            RowsCopied   = $RowsCopied
            Message      = $Result -match ".+" -notmatch "Starting copy..."
        }
    }
} # Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -Path "csv\Data.csv" -UTF8
# Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -Path "data\Data.data" -NoHeaders -UTF8
# Import-MssqlCsv "192.168.3.123,1433" "kaede" "1230" "[CHG].[CHG].[TEST]" "csv\Data.csv" -UTF8
