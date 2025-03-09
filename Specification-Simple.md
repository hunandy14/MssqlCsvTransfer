# Export-MssqlToCsv PowerShell 模組規格說明

## 1. Split-SqlTableName 函數
功能：拆分 SQL 表名稱
- 輸入：接受各種格式的表名 (如 `[DB].[Schema].[Table]`, `Schema.Table`, `Table`)
- 輸出：返回一個包含以下屬性的物件：
  - DatabaseName：資料庫名
  - SchemaName：模式名
  - TableName：表名
  - FullTableName：完整格式化的表名

## 2. Export-MssqlToCsv 函數
功能：將 SQL Server 的資料匯出為 CSV 檔案

### 主要參數組合（三種模式）
1. 表格模式 (-Table)：
   - 直接匯出指定表格
   - 例：`Export-MssqlToCsv "server" "user" "pwd" "DB.Schema.Table"`
   - 會自動獲取表格欄位資訊並生成適當的 SQL 查詢

2. SQL 檔案模式 (-SQLPath)：
   - 從 SQL 檔案讀取查詢
   - 例：`Export-MssqlToCsv "server" "user" "pwd" -SQLPath "query.sql"`
   - 輸出檔名預設使用 SQL 檔案名稱（不含副檔名）

3. SQL 查詢模式 (-SQLQuery)：
   - 直接執行 SQL 查詢字串
   - 支援管道輸入
   - 例：`Export-MssqlToCsv "server" "user" "pwd" -SQLQuery "SELECT * FROM Table"`
   - 預設輸出檔名為 'QueryResult.csv'

### 重要功能選項

#### 1. 編碼設定
- -UTF8：無 BOM 的 UTF-8
- -UTF8BOM：有 BOM 的 UTF-8
- -Encoding：自定義編碼

#### 2. 輸出控制
- -Path：指定輸出路徑
- -OutToTemp：輸出到臨時目錄
