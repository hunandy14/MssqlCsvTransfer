BEGIN TRANSACTION
SET QUOTED_IDENTIFIER ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
COMMIT
BEGIN TRANSACTION
GO
CREATE TABLE CHG.TEST
	(
	PK_ID int NOT NULL IDENTITY (1, 1),
	Name nvarchar(16) NULL,
	Mail nvarchar(64) NULL,
	Phone nvarchar(16) NULL
	)  ON [PRIMARY]
GO
ALTER TABLE CHG.TEST ADD CONSTRAINT
	PK_TEST PRIMARY KEY CLUSTERED 
	(
	PK_ID
	) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]

GO
ALTER TABLE CHG.TEST SET (LOCK_ESCALATION = TABLE)
GO
COMMIT