-- 如果表格存在就刪除
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[Table02]') AND type in (N'U'))
DROP TABLE [Table02]
GO

CREATE TABLE [Table02] (
    [Id]    INT            IDENTITY (1, 1) NOT NULL,
    [Name]  NVARCHAR (50)  COLLATE Chinese_Taiwan_Stroke_90_CI_AS NULL,
    [Value] NVARCHAR (200) COLLATE Chinese_Taiwan_Stroke_90_CI_AS NULL,
    [Date]  DATETIME2 (3)  NULL,
    CONSTRAINT [PK_Table02] PRIMARY KEY CLUSTERED ([Id] ASC)
);