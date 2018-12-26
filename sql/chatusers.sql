USE [integrator]
GO

/****** Object:  Table [dbo].[chatusers]    Script Date: 12/26/2018 12:18:14 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[chatusers](
	[chat_id] [nvarchar](50) NOT NULL,
	[SID] [nvarchar](1024) NOT NULL,
	[Registered] [bit] NOT NULL,
	[DisplayName] [nvarchar](1024) NULL
) ON [PRIMARY]
GO
