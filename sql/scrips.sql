USE [integrator]
GO

/****** Object:  Table [dbo].[scripts]    Script Date: 12/26/2018 12:34:30 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[scripts](
	[command] [nvarchar](50) NOT NULL,
	[scriptblock] [nvarchar](max) NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
