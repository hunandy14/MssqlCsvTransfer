-- 清空表格
TRUNCATE TABLE [CHG].[CHG].[Table02];
GO

-- 插入測試資料
INSERT INTO [CHG].[CHG].[Table02] ([Name], [Value], [Date])
VALUES 
    -- 基礎測試
    (N'基本數值', N'1', GETDATE()),                       -- 基本測試
    (N'NULL值', NULL, GETDATE()),                       -- NULL 值測試
    (N'空字串', N'', GETDATE()),                         -- 空字串測試
    
    -- 分隔符號測試
    (N'含逗號', N'Hello,World', GETDATE()),             -- 含逗號測試
    (N'含分號', N'Hello;World', GETDATE()),             -- 含分號測試
    (N'含TAB', N'Hello	World', GETDATE()),              -- 含 TAB 測試
    (N'含豎線', N'Hello|World', GETDATE()),             -- 含豎線測試
    (N'多逗號', N'1,2,3,4,5', GETDATE()),               -- 多個逗號測試
    (N'首逗號', N',ABC', GETDATE()),                    -- 開頭是逗號
    (N'尾逗號', N'ABC,', GETDATE()),                    -- 結尾是逗號
    
    -- 引號測試
    (N'雙引號', N'Hello"Quote"', GETDATE()),            -- 含雙引號測試
    (N'單引號', N'Hello''Quote', GETDATE()),            -- 含單引號測試
    (N'首尾雙', N'"Hello"', GETDATE()),                 -- 首尾雙引號測試
    (N'首尾單', N'''Hello''', GETDATE()),               -- 首尾單引號測試
    (N'多引號', N'"""Hello""World"""', GETDATE()),      -- 多重引號測試
    (N'引逗號', N'"Hello","World"', GETDATE()),         -- 引號加逗號測試
    
    -- 空白字元測試
    (N'純空白', N'    ', GETDATE()),                    -- 含空白字元測試
    (N'零寬度', N'Hello​World', GETDATE()),             -- 零寬度空格測試 U+200B
    (N'不斷行', N'Hello World', GETDATE()),            -- 不斷行空格測試 U+00A0
    (N'半形空', N'Hello World', GETDATE()),            -- 一般空格測試 U+0020
    (N'全形空', N'Hello　World', GETDATE()),            -- 全形空格測試 U+3000
    (N'中空格', N'Hello World', GETDATE()),            -- 中日韓空格測試 U+2002
    (N'窄空格', N'Hello World', GETDATE()),            -- 窄空格測試 U+2009
    (N'數空格', N'Hello World', GETDATE()),            -- 數學空格測試 U+205F
    (N'結尾空格', N'Hello World ', GETDATE()),         -- 結尾空格測試
    (N'結尾全形空格', N'Hello World　', GETDATE()),    -- 結尾全形空格測試
    (N'換行LF', N'Hello
World', GETDATE()),                                    -- 含換行符測試
    (N'換行CR', N'Hello\r\nWorld', GETDATE()),          -- 含 CRLF 測試
    (N'純TAB', CHAR(9), GETDATE()),                    -- 純 TAB 字元測試
    (N'多空白', REPLICATE(N' ', 10), GETDATE()),        -- 多個空白測試
    
    -- 特殊字元測試
    (N'特符號', N'©®™', GETDATE()),                     -- 特殊符號測試
    (N'中文字', N'中文測試', GETDATE()),                -- 中文測試
    (N'表情符', N'🌟🎉', GETDATE()),                    -- Emoji 測試
    (N'歸位CR', CHAR(13), GETDATE()),                  -- 純 CR 字元測試
    (N'反斜線', N'\\', GETDATE()),                      -- 反斜線測試
    (N'特殊組', N'©Hello,World®', GETDATE()),          -- 特殊符號加逗號
    (N'全形符', N'，。：', GETDATE()),                  -- 全形標點符號
    (N'混合符', N'Hello，World', GETDATE()),            -- 半形全形混合
    
    -- 極端值測試
    (N'長字串', REPLICATE(N'A', 100), GETDATE()),       -- 長字串測試
    (N'前後空', N'   Hello   ', GETDATE()),             -- 前後空白測試
    (N'零字元', CHAR(0), GETDATE()),                   -- NULL 字元測試
    (N'控制符', CHAR(1), GETDATE()),                   -- 控制字元測試
    (N'混合符', N'1,2;3	4
5|6', GETDATE()),                                      -- 混合分隔符測試
    (N'極端例', N'"""Hello,;	
World"""', GETDATE()),                                 -- 極端混合測試
    (N'數字符', N'$1,234.56', GETDATE()),              -- 數值格式測試
    (N'日期符', N'2024/03/08, 14:30', GETDATE()),      -- 日期格式測試
    (N'網址符', N'https://example.com,index.html', GETDATE()), -- URL測試
    (N'XML符', N'<tag>,</tag>', GETDATE());            -- XML標記測試
GO

-- 查看結果
SELECT * FROM [CHG].[CHG].[Table02];
GO

-- 匯出建議：
-- 1. 使用 SSMS 的 "將結果儲存為" 功能
-- 2. 選擇 CSV 格式
-- 3. 分別測試 "引號識別字" 選項的開啟和關閉狀態
-- 4. 觀察 "Unicode" 選項對中文和特殊符號的影響 
