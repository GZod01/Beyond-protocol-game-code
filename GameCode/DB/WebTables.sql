USE [BP01]
-- =============================================
-- Script Template
-- =============================================
/****** Object:  Table [dbo].[OperatorTransID]    Script Date: 04/30/2011 17:27:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
if EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[OperatorTransID]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	DROP TABLE [dbo].[OperatorTransID]
GO

CREATE TABLE [dbo].[OperatorTransID](
	[LastTransactionID] [bigint] NULL
) ON [PRIMARY]
GO

DELETE FROM OperatorTransID
INSERT INTO OperatorTransID (LastTransactionID) VALUES (0)

/****** Object:  Table [dbo].[tblWebsiteTrans]    Script Date: 04/30/2011 17:27:51 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
if EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[WebsiteTrans]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	DROP TABLE [dbo].[WebsiteTrans]
GO

CREATE TABLE [dbo].[WebsiteTrans](
	[TransactionID] [bigint] IDENTITY(1,1) NOT NULL,
	[PreviousUserName] [nvarchar](50) NULL,
	[NewUserName] [nvarchar](50) NULL,
	[NewPassword] [nvarchar](50) NULL,
	[NewStatus] [int] NULL,
	[WebUserID] [int] NULL,
	[BBUserID] [int] NULL,
	[TrackerUserID] [int] NULL,
 CONSTRAINT [PK_WebsiteTrans] PRIMARY KEY CLUSTERED 
(
	[TransactionID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

DELETE FROM WebsiteTrans


if EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[User_Play_HX]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	DROP TABLE [dbo].[User_Play_HX]
GO


CREATE TABLE [dbo].[User_Play_HX](
	[PlayerID] [int] NOT NULL,
	[UserName] [nvarchar](50) NOT NULL,
	[LoginTime] [datetime] NOT NULL,
	[IPAddress] [nvarchar](50) NOT NULL,
	[CurrentStatus] [int] NOT NULL,
	[CurrentPlayed] [int] NOT NULL,
 CONSTRAINT [PK_User_Play_HX] PRIMARY KEY CLUSTERED 
(
	[PlayerID], [UserName], [LoginTime], [IPAddress], [CurrentStatus], [CurrentPlayed] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


if EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[fc_module_users]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	DROP TABLE [dbo].[fc_module_users]
GO

CREATE TABLE [dbo].[fc_module_users](
	[ID] [bigint] IDENTITY(1,1) NOT NULL,
	[UserName] [nvarchar](50) NOT NULL,
	[GameStartDate] [datetime] NOT NULL,
	[SubExpirationDate] [datetime] NOT NULL,
 CONSTRAINT [PK_fc_module_users] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


if EXISTS (SELECT * FROM dbo.sysobjects WHERE ID = object_id(N'[dbo].[fc_subscriptions]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
	DROP TABLE [dbo].[fc_subscriptions]
GO

CREATE TABLE [dbo].[fc_subscriptions](
	[UserID] [nvarchar](50) NOT NULL,
	[UserName] [nvarchar](50) NOT NULL,
	[StartDate] [datetime] NOT NULL,
	[EndDate] [datetime] NOT NULL,
	[ProdID] [int] NOT NULL,
 CONSTRAINT [PK_fc_subscriptions] PRIMARY KEY CLUSTERED 
(
	[UserID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO




-- To add a new username
--SET IDENTITY_INSERT WebsiteTrans ON
--INSERT INTO WebsiteTrans (TransactionID, PreviousUserName, NewUserName, NewPassword, NewStatus, WebUserID, BBUserID, TrackerUserID ) VALUES (1, NULL, 'username', 'password', 1, 100, 100, 100)
--SET IDENTITY_INSERT WebsiteTrans OFF

--SET IDENTITY_INSERT fc_module_users ON
--INSERT INTO fc_module_users (ID, UserName, GameStartDate, SubExpirationDate) VALUES (100, 'username','2011/1/1 00:00', '2012/1/1 00:00')
--SET IDENTITY_INSERT fc_module_users OFF

--INSERT INTO fc_subscriptions (UserID, UserName, StartDate, EndDate, ProdID) VALUES (100, 'username','2011/1/1 00:00', '2012/1/1 00:00', 64)
