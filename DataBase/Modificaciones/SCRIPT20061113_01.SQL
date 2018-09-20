if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tallajes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Tallajes]
GO

CREATE TABLE [dbo].[Tallajes] (
	[TallajeID] [int] IDENTITY (1, 1) NOT NULL ,
	[Nombre] [Varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Estandar] [Bit] NOT NULL,
	[DescripcionT36] [Varchar] (5) NOT NULL ,
	[DescripcionT38] [Varchar] (5) NOT NULL ,
	[DescripcionT40] [Varchar] (5) NOT NULL ,
	[DescripcionT42] [Varchar] (5) NOT NULL ,
	[DescripcionT44] [Varchar] (5) NOT NULL ,
	[DescripcionT46] [Varchar] (5) NOT NULL ,
	[DescripcionT48] [Varchar] (5) NOT NULL ,
	[DescripcionT50] [Varchar] (5) NOT NULL ,
	[DescripcionT52] [Varchar] (5) NOT NULL ,
	[DescripcionT54] [Varchar] (5) NOT NULL ,
	[DescripcionT56] [Varchar] (5) NOT NULL ,
	[PermitidoT36] [Bit] NOT NULL ,
	[PermitidoT38] [Bit] NOT NULL ,
	[PermitidoT40] [Bit] NOT NULL ,
	[PermitidoT42] [Bit] NOT NULL ,
	[PermitidoT44] [Bit] NOT NULL ,
	[PermitidoT46] [Bit] NOT NULL ,
	[PermitidoT48] [Bit] NOT NULL ,
	[PermitidoT50] [Bit] NOT NULL ,
	[PermitidoT52] [Bit] NOT NULL ,
	[PermitidoT54] [Bit] NOT NULL ,
	[PermitidoT56] [Bit] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Tallajes] ADD 
	CONSTRAINT [PK_Tallajes] PRIMARY KEY  NONCLUSTERED 
	(
		[TallajeID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_Tallajes] UNIQUE  NONCLUSTERED 
	(
		[NOMBRE]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

INSERT INTO [dbo].[Tallajes]
VALUES
('Estándar',
 1,
 '36',
 '38',
 '40',
 '42',
 '44',
 '46',
 '48',
 '50',
 '52',
 '54',
 '56',
1,
1,
1,
1,
1,
1,
1,
1,
1,
1,
1
)

ALTER TABLE [dbo].[Articulos] 
	ADD TallajeID INTEGER NULL
	
GO
UPDATE [dbo].[Articulos] 
	SET TallajeID = 1
GO

ALTER TABLE [dbo].[Articulos] 
	ALTER COLUMN TallajeID INTEGER NOT NULL
GO

INSERT INTO [dbo].[Tallajes]
VALUES
('Talla única',
 0,
 'TU',
 '',
 '',
 '',
 '',
 '',
 '',
 '',
 '',
 '',
 '',
1,
0,
0,
0,
0,
0,
0,
0,
0,
0,
0
)

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vArticulos]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vArticulos]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto: vista dbo.vArticulos    fecha de la secuencia de comandos: 13/09/2001 20:09:22 *****
***** Objeto:  vista dbo.vArticulos    fecha de la secuencia de comandos: 10/09/2001 18:06:20 ******/
CREATE VIEW dbo.vArticulos
AS
SELECT     dbo.Articulos.ARTICULOID, dbo.Articulos.NOMBRE, dbo.Articulos.STOCKACTUAL, dbo.Articulos.STOCKPENDIENTE, dbo.Articulos.STOCKMINIMO, 
                      dbo.Articulos.STOCKMAXIMO, dbo.Articulos.LOTEECONOMICO, dbo.Articulos.PRECIOCOSTEEUR, dbo.Articulos.PRECIOVENTAEUR, 
                      dbo.Articulos.PRECIOVENTAPUBLICO, RTRIM(dbo.Series.Nombre) AS NOMBRESERIE, RTRIM(dbo.Prendas.NOMBRE) AS NOMBREPRENDA, 
                      RTRIM(dbo.Modelos.NOMBRE) AS NOMBREMODELO, dbo.Articulos.TEMPORADAID, dbo.Articulos.PRENDAID, dbo.Articulos.MODELOID, 
                      dbo.Articulos.SERIEID, dbo.Articulos.SuReferencia, dbo.Articulos.ProveedorID, dbo.Articulos.PrecioCompraEUR, dbo.Articulos.TallajeID, 
                      dbo.Tallajes.Nombre AS NombreTallaje
FROM         dbo.Articulos INNER JOIN
                      dbo.Series ON dbo.Articulos.SERIEID = dbo.Series.SerieID INNER JOIN
                      dbo.Prendas ON dbo.Articulos.PRENDAID = dbo.Prendas.PrendaID INNER JOIN
                      dbo.Modelos ON dbo.Articulos.MODELOID = dbo.Modelos.MODELOID INNER JOIN
                      dbo.Tallajes ON dbo.Articulos.TallajeID = dbo.Tallajes.TallajeID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

