# 移除 CSV 數據中的雙引號
function Remove-CsvQuotes {
    param (
        [Parameter(ValueFromPipeline = $true)]
        [psobject[]]$InputObject,
        [string]$InputPath,
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        [string]$Delimiter = ',',
        [System.Text.Encoding]$Encoding = (New-Object System.Text.UTF8Encoding $False),
        [switch]$RemoveHeader
    )

    begin {
        $directory = Split-Path -Path $OutputPath -Parent
        if (-not (Test-Path $directory)) { New-Item -ItemType Directory -Path $directory | Out-Null }
        $writer = New-Object System.IO.StreamWriter -ArgumentList $OutputPath, $false, $Encoding
        $headerProcessed = $false
    }

    process {
        if ($InputObject) {
            foreach ($obj in $InputObject) {
                $line = ""
                $properties = $obj | Get-Member -MemberType Properties
                foreach ($prop in $properties) {
                    if (-not [string]::IsNullOrEmpty($line)) { $line += $Delimiter }
                    $value = $obj.$($prop.Name) -replace '"', '""'
                    $line += $value
                }; $writer.WriteLine($line)
            }
        } elseif ($InputPath) {
            $reader = New-Object System.IO.StreamReader -ArgumentList $InputPath, $Encoding
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                if ($RemoveHeader -and -not $headerProcessed) {
                    $headerProcessed = $true
                    continue
                }
                $fields = $line.Split($Delimiter)
                $newLine = ""
                for ($i = 0; $i -lt $fields.Length; $i++) {
                    $cleanField = $fields[$i].Trim('"')
                    if ($i -gt 0) { $newLine += $Delimiter }
                    $newLine += $cleanField
                }; $writer.WriteLine($newLine)
            }; $reader.Close()
        }
    }

    end {
        $writer.Close()
    }
} # Remove-CsvQuotes -InputPath 'csv\Data.csv' -OutputPath 'data\Data.data' -RemoveHeader



# 上傳 CSV 到 MSSQL資料庫
function Import-MssqlCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $ServerName,
        [Parameter(Mandatory)]
        [string] $UserName,
        [Parameter(Mandatory)]
        [string] $Passwd,
        [Parameter(Mandatory)]
        [string] $Table,
        [Parameter(Mandatory)]
        [string] $CsvPath,
        [string] $Terminator = ',',
        [string] $RowTerminator = "`n",
        [switch] $OutNull
    )
    
    begin {
        $dataPath = "data\Data.data"
        Remove-CsvQuotes -InputPath $CsvPath -OutputPath $dataPath -RemoveHeader
        $CsvPath = $dataPath
    }
    
    process {
        $Output = & bcp $Table in $CsvPath -c -t $Terminator -r $RowTerminator -S $ServerName -U $UserName -P $Passwd
        $HasError = $false
        $RowsCopied = 0
        
        $OutputString = $Output -join "`r`n"
        if ($outputString -match "(\d+) rows copied\.") {
            $RowsCopied = [int]$matches[1]
            if ($RowsCopied -eq 0) { $HasError = $true }
        } else { $HasError = $true }
        
        if (-not $OutNull) {
            if ($HasError) {
                $ErrMsg = $Output -join ', '
                Write-Error "BCP 命令執行失敗:: $ErrMsg"
            } else { Write-Host "BCP 命令執行成功, 共複製了 $RowsCopied 行" }
        }
    }
    
    end {
        return [pscustomobject]@{
            Success = !$HasError
            RowsCopied = $RowsCopied
            Message = $Output -match ".+" -notmatch "Starting copy..."
        }
    }
} # Import-MssqlCsv -ServerName "192.168.3.123,1433" -UserName "kaede" -Passwd "1230" -Table "[CHG].[CHG].[TEST]" -CsvPath "csv\Data.csv" | Out-Null
