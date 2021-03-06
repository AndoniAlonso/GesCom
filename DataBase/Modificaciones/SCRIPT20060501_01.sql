DROP TABLE [dbo].[PedidoCompraArticulos]

CREATE TABLE [dbo].[PedidoCompraArticulos] (
	[PedidoCompraArticuloID] [int] IDENTITY (1, 1) NOT NULL ,
	[PedidoCompraID] [int] NOT NULL ,
	[ArticuloColorID] [int] NOT NULL ,
	[Situacion] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CantidadT36] [CANTIDAD] NOT NULL ,
	[CantidadT38] [CANTIDAD] NOT NULL ,
	[CantidadT40] [CANTIDAD] NOT NULL ,
	[CantidadT42] [CANTIDAD] NOT NULL ,
	[CantidadT44] [CANTIDAD] NOT NULL ,
	[CantidadT46] [CANTIDAD] NOT NULL ,
	[CantidadT48] [CANTIDAD] NOT NULL ,
	[CantidadT50] [CANTIDAD] NOT NULL ,
	[CantidadT52] [CANTIDAD] NOT NULL ,
	[CantidadT54] [CANTIDAD] NOT NULL ,
	[CantidadT56] [CANTIDAD] NOT NULL ,
	[ServidoT36] [CANTIDAD] NOT NULL ,
	[ServidoT38] [CANTIDAD] NOT NULL ,
	[ServidoT40] [CANTIDAD] NOT NULL ,
	[ServidoT42] [CANTIDAD] NOT NULL ,
	[ServidoT44] [CANTIDAD] NOT NULL ,
	[ServidoT46] [CANTIDAD] NOT NULL ,
	[ServidoT48] [CANTIDAD] NOT NULL ,
	[ServidoT50] [CANTIDAD] NOT NULL ,
	[ServidoT52] [CANTIDAD] NOT NULL ,
	[ServidoT54] [CANTIDAD] NOT NULL ,
	[ServidoT56] [CANTIDAD] NOT NULL ,
	[PrecioCompraEUR] [IMPORTEEUR] NOT NULL ,
	[Descuento] [PORCENTAJE] NOT NULL ,
	[BrutoEUR] [IMPORTEEUR] NOT NULL ,
	[Comision] [PORCENTAJE] NOT NULL ,
	[TemporadaID] [int] NOT NULL ,
	[Observaciones] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[PedidoCompraArticulos] ADD 
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT36] DEFAULT (0) FOR [CantidadT36],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT38] DEFAULT (0) FOR [CantidadT38],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT40] DEFAULT (0) FOR [CantidadT40],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT42] DEFAULT (0) FOR [CantidadT42],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT44] DEFAULT (0) FOR [CantidadT44],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT46] DEFAULT (0) FOR [CantidadT46],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT48] DEFAULT (0) FOR [CantidadT48],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT50] DEFAULT (0) FOR [CantidadT50],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT52] DEFAULT (0) FOR [CantidadT52],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT54] DEFAULT (0) FOR [CantidadT54],
	CONSTRAINT [DF_PedidoCompraArticulo_CantidadT56] DEFAULT (0) FOR [CantidadT56],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT36] DEFAULT (0) FOR [ServidoT36],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT38] DEFAULT (0) FOR [ServidoT38],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT40] DEFAULT (0) FOR [ServidoT40],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT42] DEFAULT (0) FOR [ServidoT42],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT44] DEFAULT (0) FOR [ServidoT44],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT46] DEFAULT (0) FOR [ServidoT46],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT48] DEFAULT (0) FOR [ServidoT48],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT50] DEFAULT (0) FOR [ServidoT50],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT52] DEFAULT (0) FOR [ServidoT52],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT54] DEFAULT (0) FOR [ServidoT54],
	CONSTRAINT [DF_PedidoCompraArticulo_ServidoT56] DEFAULT (0) FOR [ServidoT56],
	CONSTRAINT [DF_PedidoCompraArticulos_PrecioCompraEUR] DEFAULT (0) FOR [PrecioCompraEUR],
	CONSTRAINT [DF_PedidoCompraArticulo_Descuento] DEFAULT (0) FOR [Descuento],
	CONSTRAINT [DF_PedidoCompraArticulo_BrutoEUR] DEFAULT (0) FOR [BrutoEUR],
	CONSTRAINT [DF_PedidoCompraArticulo_Comision] DEFAULT (0) FOR [Comision],
	CONSTRAINT [PK_PedidoCompraArticulos] PRIMARY KEY  NONCLUSTERED 
	(
		[PedidoCompraArticuloID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[PedidoCompraArticulos] ADD 
	CONSTRAINT [FK_PedidoCompraArticulos_ArticuloColores] FOREIGN KEY 
	(
		[ArticuloColorID]
	) REFERENCES [dbo].[ArticuloColores] (
		[ARTICULOCOLORID]
	),
	CONSTRAINT [FK_PedidoCompraArticulos_PedidosCompra] FOREIGN KEY 
	(
		[PedidoCompraID]
	) REFERENCES [dbo].[PedidosCompra] (
		[PedidoCompraID]
	)
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ArticuloColorAlmacen]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ArticuloColorAlmacen]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Almacenes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Almacenes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CentrosGestion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CentrosGestion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[0]') and OBJECTPROPERTY(id, N'IsDefault') = 1)
drop default [dbo].[0]
GO

create default [0] as 0

GO


CREATE TABLE [dbo].[CentrosGestion] (
	[CentroGestionID] [int] IDENTITY (1,1) NOT NULL ,
	[Nombre] [varchar] (100) NOT NULL ,
	[DireccionID] [int] NOT NULL ,
	[ContadorTicketID] [int] NULL ,
	[ContadorPedidoVentaID] [int] NULL ,
	[ContadorAlbaranVentaID] [int] NULL ,
	[ContadorFacturaVentaID] [int] NULL ,
	[ContadorPedidoCompraID] [int] NULL ,
	[ContadorAlbaranCompraID] [int] NULL ,
	[ContadorFacturaCompraID] [int] NULL ,
	[SedeCentral] [bit] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[CentrosGestion] WITH NOCHECK ADD 
	CONSTRAINT [PK_CentrosGestion] PRIMARY KEY  CLUSTERED 
	(
		[CentroGestionID]
	)  ON [PRIMARY] 
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[0]', N'[CentrosGestion].[SedeCentral]'
GO

setuser
GO

ALTER TABLE [dbo].[CentrosGestion] ADD 
	CONSTRAINT [FK_CentrosGestion_Direcciones] FOREIGN KEY 
	(
		[DireccionID]
	) REFERENCES [dbo].[Direcciones] (
		[DireccionID]
	)
GO
INSERT INTO CentrosGestion (
	[Nombre],
	[DireccionID],
	[ContadorTicketID],
	[ContadorPedidoVentaID],
	[ContadorAlbaranVentaID],
	[ContadorFacturaVentaID],
	[ContadorPedidoCompraID],
	[ContadorAlbaranCompraID],
	[ContadorFacturaCompraID],
	[SedeCentral]) VALUES 
	('HONGO CENTRAL' ,
	 1,
	 NULL,
	 NULL,
	 NULL,
	 NULL,
	 NULL,
	 NULL,
	 NULL,
	 1
         ) 
GO


CREATE TABLE [dbo].[Almacenes] (
	[AlmacenID] [int] IDENTITY (1,1) NOT NULL ,
	[Nombre] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CentroGestionID] [int] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Almacenes] WITH NOCHECK ADD 
	CONSTRAINT [PK_Almacenes] PRIMARY KEY  CLUSTERED 
	(
		[AlmacenID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Almacenes] ADD 
	CONSTRAINT [FK_Almacenes_CentrosGestion] FOREIGN KEY 
	(
		[CentroGestionID]
	) REFERENCES [dbo].[CentrosGestion] (
		[CentroGestionID]
	)
GO

INSERT INTO Almacenes values ('Central', 1)
GO

CREATE TABLE [dbo].[ArticuloColorAlmacen] (
	[ArticuloColorID] [int] NOT NULL ,
	[AlmacenID] [int] NOT NULL ,
	[STOCKACTUALT36] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT38] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT40] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT42] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT44] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT46] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT48] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT50] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT52] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT54] [CANTIDAD] NOT NULL ,
	[STOCKACTUALT56] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET36] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET38] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET40] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET42] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET44] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET46] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET48] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET50] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET52] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET54] [CANTIDAD] NOT NULL ,
	[STOCKPENDIENTET56] [CANTIDAD] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[ArticuloColorAlmacen] WITH NOCHECK ADD 
	CONSTRAINT [PK_ArticuloColorAlmacen] PRIMARY KEY  CLUSTERED 
	(
		[ArticuloColorID],
		[AlmacenID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[ArticuloColorAlmacen] ADD 
	CONSTRAINT [FK_ArticuloColorAlmacen_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	),
	CONSTRAINT [FK_ArticuloColorAlmacen_ArticuloColores] FOREIGN KEY 
	(
		[ArticuloColorID]
	) REFERENCES [dbo].[ArticuloColores] (
		[ARTICULOCOLORID]
	)
GO

--ALTER TABLE PedidosCompra DROP CONSTRAINT [FK_PedidosCompra_Almacenes]
--GO
--ALTER TABLE PedidosCompra DROP COLUMN AlmacenID
--GO
--ALTER TABLE PedidosCompra
--	ADD AlmacenID int
--GO
--UPDATE PedidosCompra
--	SET AlmacenID = 1
--GO
--ALTER TABLE PedidosCompra
--	ALTER COLUMN AlmacenID int NOT NULL
--GO
--ALTER TABLE [dbo].[PedidosCompra] ADD 
--	CONSTRAINT [FK_PedidosCompra_Almacenes] FOREIGN KEY 
--	(
--		[AlmacenID]
--	) REFERENCES [dbo].[Almacenes] (
--		[AlmacenID]
--	)
--GO



GO
SET ANSI_NULLS ON 
GO

ALTER TABLE MoviArticulos DROP CONSTRAINT [FK_MoviArticulos_Almacenes]
GO
ALTER TABLE MoviArticulos DROP COLUMN AlmacenID
GO
ALTER TABLE MoviArticulos 
	ADD AlmacenID int
GO
--UPDATE MoviArticulos
--	SET AlmacenID = 1
--GO
--ALTER TABLE MoviArticulos 
--	ALTER COLUMN AlmacenID int NOT NULL
--go
ALTER TABLE [dbo].[MoviArticulos] ADD 
	CONSTRAINT [FK_MoviArticulos_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	)
GO

ALTER TABLE PedidoCompraArticulos DROP CONSTRAINT [FK_PedidoCompraArticulos_Almacenes]
GO
ALTER TABLE PedidoCompraArticulos DROP COLUMN AlmacenID
GO
ALTER TABLE PedidoCompraArticulos 
	ADD AlmacenID int
GO
UPDATE PedidoCompraArticulos
	SET AlmacenID = 1
GO
ALTER TABLE PedidoCompraArticulos 
	ALTER COLUMN AlmacenID int NOT NULL
go
ALTER TABLE [dbo].[PedidoCompraArticulos] ADD 
	CONSTRAINT [FK_PedidoCompraArticulos_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	)
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoCompraArticulos]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoCompraArticulos]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoCompraArticulos
AS
SELECT     dbo.PedidoCompraArticulos.*, dbo.vNombreArticuloColores.Nombre AS NombreArticuloColor
FROM         dbo.PedidoCompraArticulos INNER JOIN
                      dbo.vNombreArticuloColores ON dbo.PedidoCompraArticulos.ArticuloColorID = dbo.vNombreArticuloColores.ARTICULOCOLORID

GO

DROP TABLE [dbo].[ParametrosAplicacion]
GO
CREATE TABLE [dbo].[ParametrosAplicacion] (
	[ParametroAplicacionID] [char] (10)  NOT NULL ,
	[Nombre] [varchar] (100)  NULL ,
	[Valor] [varchar] (100)  NULL ,
	[Sistema] [bit] NOT NULL ,
	[TipoParametro] [int] NULL ,
	CONSTRAINT [PK_ParametrosAplicacion] PRIMARY KEY  CLUSTERED 
	(
		[ParametroAplicacionID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY]
)
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[0]', N'[ParametrosAplicacion].[Sistema]'
GO

setuser
GO

INSERT INTO ParametrosAplicacion values ('ALMAPRED', 'Almacén predeterminado', '1', 1, 1)
GO

