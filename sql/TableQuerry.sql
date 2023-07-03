-- 変数の宣言
DECLARE @Table NVARCHAR(128) = '[CHG].[CHG].[Employees]'

-- 完全なテーブル名から短いテーブル名を抽出する
DECLARE @TableName NVARCHAR(128)
SET @TableName = (SELECT REVERSE(LEFT(REVERSE(@Table), CHARINDEX('.', REVERSE(@Table)) - 1)))
SET @TableName = REPLACE(REPLACE(@TableName, '[', ''), ']', '')

-- カラム名を含むSQL文を生成する
DECLARE @Columns NVARCHAR(MAX)
SELECT @Columns = STRING_AGG(CAST(('''"' + COLUMN_NAME + '"''')  AS NVARCHAR(MAX)), ', ')
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = @TableName;

-- データを含むSQL文を生成する
DECLARE @Data NVARCHAR(MAX)
SELECT @Data = STRING_AGG(CAST((' ''"'' + REPLACE(ISNULL(' + COLUMN_NAME + ',''''), ''"'', ''""'') + ''"'' AS ' + COLUMN_NAME) AS NVARCHAR(MAX)), ', ')
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = @TableName;

-- 最終的なSQLクエリ文を組み立てる
DECLARE @SQL NVARCHAR(MAX)
SET @SQL = 'SET NOCOUNT ON;' + '
SELECT 
  '+ @Columns + ';' + '
SELECT
  ' + @Data + '
FROM
  ' + @Table + ';';

-- 動的SQLを実行する
EXEC sp_executesql @SQL
