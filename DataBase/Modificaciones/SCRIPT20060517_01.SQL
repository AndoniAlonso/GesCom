ALTER TABLE [dbo].[Articulos] 
	ADD SuReferencia CHAR (30) NULL,
	    ProveedorID INT NULL
GO


ALTER TABLE [dbo].[Articulos] 
	ADD 
	CONSTRAINT [FK_Articulos_Proveedores] FOREIGN KEY 
	(
		[ProveedorID]
	) REFERENCES [dbo].[Proveedores] (
		[ProveedorID]
	)
GO
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
                      dbo.Articulos.SERIEID, dbo.Articulos.SuReferencia, dbo.Articulos.ProveedorID
FROM         dbo.Articulos INNER JOIN
                      dbo.Series ON dbo.Articulos.SERIEID = dbo.Series.SerieID INNER JOIN
                      dbo.Prendas ON dbo.Articulos.PRENDAID = dbo.Prendas.PrendaID INNER JOIN
                      dbo.Modelos ON dbo.Articulos.MODELOID = dbo.Modelos.MODELOID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

INSERT INTO ParametrosAplicacion values ('TIPOINST', 'Tipo instalacion (0-->Fabrica, 1--> punto de venta', '0', 1, 1)
GO

INSERT INTO Materiales VALUES ( 
	'Material predeterminado', --NOMBRE
	'MATEPRED', -- CODIGO 
	'U',  --UNIDADES
	0, --[STOCKACTUAL] [CANTIDAD] NOT NULL ,
	0, --[STOCKPENDIENTE] [CANTIDAD] NOT NULL ,
	0, --[STOCKMINIMO] [CANTIDAD] NOT NULL ,
	0, --[STOCKMAXIMO] [CANTIDAD] NOT NULL ,
	0, --[PRECIOCOSTEPTA] [IMPORTEPTA] NOT NULL ,
	0, --[PRECIOCOSTEEUR] [PRECIOMATERIALEUR] NOT NULL ,
	0, --[PRECIOPONDERADOPTA] [IMPORTEPTA] NOT NULL ,
	0, --[PRECIOPONDERADOEUR] [PRECIOMATERIALEUR] NOT NULL ,
	'O', --[TipoMaterial] [char] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	1.5, --[AnchuraTela] [CANTIDAD] NOT NULL ,
	GETDATE(), --[FechaAlta] [datetime] NOT NULL ,
	' ', --[Composicion1] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	0, --[PorcComposicion1] [PORCENTAJE] NOT NULL ,
	' ', --[Composicion2] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	0, --[PorcComposicion2] [PORCENTAJE] NOT NULL ,
	' ', --[Composicion3] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	0, --[PorcComposicion3] [PORCENTAJE] NOT NULL ,
	' ', --[Composicion4] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	0, --[PorcComposicion4] [PORCENTAJE] NOT NULL ,
	NULL) --[ProveedorID] [int] NULL 
GO
DECLARE @MaterialID AS INTEGER
SELECT @MaterialID=MAX(MaterialID) FROM MATERIALES
PRINT 'MaterialID'
PRINT @MaterialID
INSERT INTO ParametrosAplicacion values ('MATEPRED', 'Material predeterminado', @MaterialID, 1, 1)
GO



if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AlbaranCompraArticulos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AlbaranCompraArticulos]
GO

CREATE TABLE [dbo].[AlbaranCompraArticulos] (
	[AlbaranCompraArticuloID] [int] IDENTITY (1, 1) NOT NULL ,
	[AlbaranCompraID] [int] NOT NULL ,
	[ArticuloColorID] [int] NULL ,
	[Descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[PedidoCompraArticuloID] [int] NULL ,
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
	[PrecioCompraEUR] [IMPORTEEUR] NOT NULL ,
	[Descuento] [PORCENTAJE] NOT NULL ,
	[BrutoEUR] [IMPORTEEUR] NOT NULL ,
	[Comision] [PORCENTAJE] NOT NULL ,
	[TemporadaID] [int] NOT NULL ,
	[FacturaCompraArticuloID] [int] NULL ,
	[AlmacenID] [int] NOT NULL ,
	CONSTRAINT [FK_AlbaranCompraArticulos_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	),
	CONSTRAINT [PK_AlbaranCompraArticulos] PRIMARY KEY  NONCLUSTERED 
	(
		[AlbaranCompraArticuloID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [FK_AlbaranCompraArticulos_AlbaranesCompra] FOREIGN KEY 
	(
		[AlbaranCompraID]
	) REFERENCES [dbo].[AlbaranesCompra] (
		[AlbaranCompraID]
	),
	CONSTRAINT [FK_AlbaranCompraArticulos_ArticuloColores] FOREIGN KEY 
	(
		[ArticuloColorID]
	) REFERENCES [dbo].[ArticuloColores] (
		[ARTICULOCOLORID]
	),
--	CONSTRAINT [FK_AlbaranCompraArticulos_FacturaCompraArticulos] FOREIGN KEY 
--	(
--		[FacturaCompraItemID]
--	) REFERENCES [dbo].[FacturaCompraArticulos] (
--		[FacturaCompraItemID]
--	),
--OJOOJO: CUANDO ESTÉ LA TABLA DE FACTURAS CREADA SE DEBE ACTIVAR ESTA FOREIGN KEY
	CONSTRAINT [FK_AlbaranCompraArticulos_PedidoCompraArticulos] FOREIGN KEY 
	(
		[PedidoCompraArticuloID]
	) REFERENCES [dbo].[PedidoCompraArticulos] (
		[PedidoCompraArticuloID]
	)
) ON [PRIMARY]
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vCobrosPendientes]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vCobrosPendientes]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vCobrosPendientes
AS
SELECT     dbo.CobrosPagos.SituacionComercial, dbo.CobrosPagos.CobroPagoID AS DocumentoID, dbo.FacturasVenta.Numero, dbo.Clientes.Nombre, 
                      'Vcto: ' + CONVERT(char, dbo.CobrosPagos.Vencimiento, 103) AS Descripcion, dbo.CobrosPagos.ImporteEUR AS Cantidad, 
                      dbo.CobrosPagos.FechaEmision, dbo.CobrosPagos.PersonaID AS PropietarioID, dbo.CobrosPagos.TemporadaID, dbo.CobrosPagos.EmpresaID, 
                      dbo.FormasDePago.MesesPrimerGiro, dbo.FormasDePago.MesesEntreGiros, dbo.CobrosPagos.Vencimiento AS Fecha
FROM         dbo.Clientes INNER JOIN
                      dbo.CobrosPagos ON dbo.Clientes.ClienteID = dbo.CobrosPagos.PersonaID INNER JOIN
                      dbo.FormasDePago ON dbo.CobrosPagos.FormaPagoID = dbo.FormasDePago.FormaPagoID LEFT OUTER JOIN
                      dbo.FacturasVenta ON dbo.CobrosPagos.FacturaID = dbo.FacturasVenta.FacturaVentaID
WHERE     (dbo.FormasDePago.MesesPrimerGiro <> 0) AND (dbo.CobrosPagos.Tipo = 'C') AND (dbo.CobrosPagos.SituacionComercial = 'A') OR
                      (dbo.CobrosPagos.Tipo = 'C') AND (dbo.CobrosPagos.SituacionComercial = 'A') AND (dbo.FormasDePago.MesesEntreGiros <> 0)

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidoCompraArticuloPendientes]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidoCompraArticuloPendientes]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidoCompraArticuloPendientes
AS
SELECT     dbo.PedidoCompraArticulos.Situacion, dbo.PedidoCompraArticulos.PedidoCompraArticuloID AS DocumentoID, dbo.PedidosCompra.Numero, 
                      dbo.PedidoCompraArticulos.CantidadT36 + dbo.PedidoCompraArticulos.CantidadT38 + dbo.PedidoCompraArticulos.CantidadT40 + dbo.PedidoCompraArticulos.CantidadT42
                       + dbo.PedidoCompraArticulos.CantidadT44 + dbo.PedidoCompraArticulos.CantidadT46 + dbo.PedidoCompraArticulos.CantidadT48 + dbo.PedidoCompraArticulos.CantidadT50
                       + dbo.PedidoCompraArticulos.CantidadT52 + dbo.PedidoCompraArticulos.CantidadT54 + dbo.PedidoCompraArticulos.CantidadT56 - dbo.PedidoCompraArticulos.ServidoT36
                       - dbo.PedidoCompraArticulos.ServidoT38 - dbo.PedidoCompraArticulos.ServidoT40 - dbo.PedidoCompraArticulos.ServidoT42 - dbo.PedidoCompraArticulos.ServidoT44
                       - dbo.PedidoCompraArticulos.ServidoT46 - dbo.PedidoCompraArticulos.ServidoT48 - dbo.PedidoCompraArticulos.ServidoT50 - dbo.PedidoCompraArticulos.ServidoT52
                       - dbo.PedidoCompraArticulos.ServidoT54 - dbo.PedidoCompraArticulos.ServidoT56 AS Cantidad, dbo.PedidosCompra.Fecha, 
                      dbo.PedidosCompra.ProveedorID AS PropietarioID, dbo.PedidosCompra.TemporadaID, dbo.PedidosCompra.EmpresaID, 
                      dbo.vNombreArticuloColores.Codigo AS Nombre, dbo.vNombreArticuloColores.Nombre AS Descripcion, 
                      dbo.PedidoCompraArticulos.ArticuloColorID
FROM         dbo.PedidosCompra INNER JOIN
                      dbo.PedidoCompraArticulos ON dbo.PedidosCompra.PedidoCompraID = dbo.PedidoCompraArticulos.PedidoCompraID INNER JOIN
                      dbo.vNombreArticuloColores ON dbo.PedidoCompraArticulos.ArticuloColorID = dbo.vNombreArticuloColores.ARTICULOCOLORID
WHERE     (dbo.PedidoCompraArticulos.Situacion <> 'C')

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

INSERT INTO Direcciones VALUES ('','', '','','','','','','','')

DECLARE @DireccionFiscalID AS INTEGER
DECLARE @DireccionEntregaID AS INTEGER
DECLARE @CuentaBancariaID AS INTEGER
DECLARE @DatoComercialAID AS INTEGER
DECLARE @DatoComercialBID AS INTEGER

SELECT @DireccionFiscalID=MAX(DireccionID)
FROM Direcciones
	
INSERT INTO Direcciones VALUES ('','', '','','','','','','','')

SELECT @DireccionEntregaID=MAX(DireccionID)
FROM Direcciones

INSERT INTO CuentasBancarias VALUES (0,0,0,0,'','')

SELECT @CuentaBancariaID=MAX(CuentaBancariaID)
FROM CuentasBancarias

INSERT INTO DatosComerciales VALUES (0,0,16)
SELECT @DatoComercialAID=MAX(DatoComercialID)
FROM DatosComerciales

INSERT INTO DatosComerciales VALUES (0,0,16)
SELECT @DatoComercialBID=MAX(DatoComercialID)
FROM DatosComerciales


INSERT INTO Clientes VALUES ( 'Clientes TPV', '', '', '', @DireccionFiscalID, @DireccionEntregaID, 1,6, @CuentaBancariaID, 1,0,0,0,'4300000001',@DatoComercialAID,@DatoComercialBID,100)

DECLARE @ClienteID AS INTEGER
SELECT @ClienteID=MAX(ClienteID)
FROM Clientes

INSERT INTO ParametrosAplicacion values ('CLIEPRED', 'Cliente predeterminado para TPV', @ClienteID, 1, 1)
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Terminales_CentrosGestion]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Terminales] DROP CONSTRAINT FK_Terminales_CentrosGestion
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Terminales]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Terminales]
GO

CREATE TABLE [dbo].[Terminales] (
	[TerminalID] [int] IDENTITY (1, 1) NOT NULL ,
	[Nombre] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CentroGestionID] [int] NOT NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Terminales] WITH NOCHECK ADD 
	CONSTRAINT [PK_Terminales] PRIMARY KEY  CLUSTERED 
	(
		[TerminalID]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

setuser
GO

ALTER TABLE [dbo].[Terminales] ADD 
	CONSTRAINT [FK_Terminales_CentrosGestion] FOREIGN KEY 
	(
		[CentroGestionID]
	) REFERENCES [dbo].[CentrosGestion] (
		[CentroGestionID]
	)
GO

-- Incluir el almacen, el terminal y el centro de gestión
ALTER TABLE [dbo].[AlbaranesVenta] 
	ADD [AlmacenID] [int] NULL ,
	    [CentroGestionID] [int] NULL ,
	    [TerminalID] [int] NULL 
GO
ALTER TABLE [dbo].[AlbaranesVenta] WITH NOCHECK ADD 
	CONSTRAINT [FK_AlbaranesVenta_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	),
	CONSTRAINT [FK_AlbaranesVenta_CentrosGestion] FOREIGN KEY 
	(
		[CentroGestionID]
	) REFERENCES [dbo].[CentrosGestion] (
		[CentroGestionID]
	),
	CONSTRAINT [FK_AlbaranesVenta_Terminales] FOREIGN KEY 
	(
		[TerminalID]
	) REFERENCES [dbo].[Terminales] (
		[TerminalID]
	)
GO
ALTER TABLE [dbo].[FacturasVenta] 
	ADD [AlmacenID] [int] NULL ,
	    [CentroGestionID] [int] NULL ,
	    [TerminalID] [int] NULL 
GO
ALTER TABLE [dbo].[FacturasVenta] WITH NOCHECK ADD 
	CONSTRAINT [FK_FacturasVenta_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	),
	CONSTRAINT [FK_FacturasVenta_CentrosGestion] FOREIGN KEY 
	(
		[CentroGestionID]
	) REFERENCES [dbo].[CentrosGestion] (
		[CentroGestionID]
	),
	CONSTRAINT [FK_FacturasVenta_Terminales] FOREIGN KEY 
	(
		[TerminalID]
	) REFERENCES [dbo].[Terminales] (
		[TerminalID]
	)
GO

INSERT INTO TERMINALES VALUES ('TPV SACHER', 1)
GO


ALTER TABLE [dbo].[FacturaVentaItems] 
	ADD [AlmacenID] [int] NULL 
GO
ALTER TABLE [dbo].[FacturaVentaItems] WITH NOCHECK ADD 
	CONSTRAINT [FK_FacturaVentaItems_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	)
GO

ALTER TABLE [dbo].[AlbaranVentaItems] 
	ADD [AlmacenID] [int] NULL 
GO
ALTER TABLE [dbo].[AlbaranVentaItems] WITH NOCHECK ADD 
	CONSTRAINT [FK_AlbaranVentaItems_Almacenes] FOREIGN KEY 
	(
		[AlmacenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	)
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTerminales]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTerminales]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTerminales
AS
SELECT     dbo.Terminales.TerminalID, dbo.Terminales.Nombre, dbo.Terminales.CentroGestionID, dbo.Almacenes.AlmacenID
FROM         dbo.Almacenes INNER JOIN
                      dbo.CentrosGestion ON dbo.Almacenes.CentroGestionID = dbo.CentrosGestion.CentroGestionID RIGHT OUTER JOIN
                      dbo.Terminales ON dbo.CentrosGestion.CentroGestionID = dbo.Terminales.CentroGestionID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Traspasos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Traspasos]
GO

CREATE TABLE [dbo].[Traspasos] (
	[TraspasoID] [int] IDENTITY (1, 1) NOT NULL ,
	[AlmacenOrigenID] [int] NOT NULL ,
	[AlmacenDestinoID] [int] NOT NULL ,
	[Situacion] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FechaAlta] [datetime] NOT NULL ,
	[FechaTransito] [datetime] NULL ,
	[FechaRecepcion] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Traspasos] WITH NOCHECK ADD 
	CONSTRAINT [PK_Traspasos] PRIMARY KEY  CLUSTERED 
	(
		[TraspasoID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Traspasos] ADD 
	CONSTRAINT [FK_Traspasos_Almacenes] FOREIGN KEY 
	(
		[AlmacenOrigenID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	),
	CONSTRAINT [FK_Traspasos_Almacenes1] FOREIGN KEY 
	(
		[AlmacenDestinoID]
	) REFERENCES [dbo].[Almacenes] (
		[AlmacenID]
	)
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TraspasoItems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TraspasoItems]
GO

CREATE TABLE [dbo].[TraspasoItems] (
	[TraspasoItemID] [int] IDENTITY (1, 1) NOT NULL ,
	[TraspasoID] [int] NOT NULL ,
	[ArticuloColorID] [int] NOT NULL ,
	[Situacion] [char] (2) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
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
	[Observaciones] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[TraspasoItems] WITH NOCHECK ADD 
	CONSTRAINT [PK_TraspasoItems] PRIMARY KEY  CLUSTERED 
	(
		[TraspasoItemID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[TraspasoItems] ADD 
	CONSTRAINT [DF_TraspasoItems_CantidadT36] DEFAULT (0) FOR [CantidadT36],
	CONSTRAINT [DF_TraspasoItems_CantidadT38] DEFAULT (0) FOR [CantidadT38],
	CONSTRAINT [DF_TraspasoItems_CantidadT40] DEFAULT (0) FOR [CantidadT40],
	CONSTRAINT [DF_TraspasoItems_CantidadT42] DEFAULT (0) FOR [CantidadT42],
	CONSTRAINT [DF_TraspasoItems_CantidadT44] DEFAULT (0) FOR [CantidadT44],
	CONSTRAINT [DF_TraspasoItems_CantidadT46] DEFAULT (0) FOR [CantidadT46],
	CONSTRAINT [DF_TraspasoItems_CantidadT48] DEFAULT (0) FOR [CantidadT48],
	CONSTRAINT [DF_TraspasoItems_CantidadT50] DEFAULT (0) FOR [CantidadT50],
	CONSTRAINT [DF_TraspasoItems_CantidadT52] DEFAULT (0) FOR [CantidadT52],
	CONSTRAINT [DF_TraspasoItems_CantidadT54] DEFAULT (0) FOR [CantidadT54],
	CONSTRAINT [DF_TraspasoItems_CantidadT56] DEFAULT (0) FOR [CantidadT56]
GO

ALTER TABLE [dbo].[TraspasoItems] ADD 
	CONSTRAINT [FK_TraspasoItems_ArticuloColores] FOREIGN KEY 
	(
		[ArticuloColorID]
	) REFERENCES [dbo].[ArticuloColores] (
		[ARTICULOCOLORID]
	),
	CONSTRAINT [FK_TraspasoItems_Traspasos] FOREIGN KEY 
	(
		[TraspasoID]
	) REFERENCES [dbo].[Traspasos] (
		[TraspasoID]
	)
GO


INSERT INTO Almacenes values ('Almacen en Transito', 1)
GO

INSERT INTO ParametrosAplicacion values ('ALMATRAN', 'Almacen en transito', '2', 1, 1)
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTraspasos]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTraspasos]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTraspasos
AS
SELECT     dbo.Traspasos.TraspasoID, dbo.Traspasos.AlmacenOrigenID, AlmacenesOrigen.Nombre AS NombreAlmacenOrigen, 
                      AlmacenesDestino.Nombre AS NombreAlmacenDestino, AlmacenesOrigen.CentroGestionID AS CentroGestionOrigenID, 
                      AlmacenesDestino.CentroGestionID AS CentroGestionDestinoID, dbo.Traspasos.Situacion, dbo.Traspasos.FechaAlta, dbo.Traspasos.FechaTransito, 
                      dbo.Traspasos.FechaRecepcion, 
                      SUM(dbo.TraspasoItems.CantidadT36 + dbo.TraspasoItems.CantidadT38 + dbo.TraspasoItems.CantidadT40 + dbo.TraspasoItems.CantidadT42 + dbo.TraspasoItems.CantidadT44
                       + dbo.TraspasoItems.CantidadT46 + dbo.TraspasoItems.CantidadT48 + dbo.TraspasoItems.CantidadT50 + dbo.TraspasoItems.CantidadT52 + dbo.TraspasoItems.CantidadT54
                       + dbo.TraspasoItems.CantidadT56) AS Cantidad
FROM         dbo.Almacenes AlmacenesOrigen INNER JOIN
                      dbo.Traspasos ON AlmacenesOrigen.AlmacenID = dbo.Traspasos.AlmacenOrigenID INNER JOIN
                      dbo.Almacenes AlmacenesDestino ON dbo.Traspasos.AlmacenDestinoID = AlmacenesDestino.AlmacenID LEFT OUTER JOIN
                      dbo.TraspasoItems ON dbo.Traspasos.TraspasoID = dbo.TraspasoItems.TraspasoID
GROUP BY dbo.Traspasos.TraspasoID, dbo.Traspasos.AlmacenOrigenID, AlmacenesOrigen.Nombre, AlmacenesDestino.Nombre, 
                      AlmacenesOrigen.CentroGestionID, AlmacenesDestino.CentroGestionID, dbo.Traspasos.Situacion, dbo.Traspasos.FechaAlta, 
                      dbo.Traspasos.FechaTransito, dbo.Traspasos.FechaRecepcion

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vNecesidadesMaterial2]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vNecesidadesMaterial2]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vNecesidadesMaterial2
AS
SELECT     dbo.PedidoVentaItems.PedidoVentaID, dbo.PedidoVentaItems.PedidoVentaItemID, dbo.Materiales.MATERIALID, dbo.Materiales.CODIGO, 
                      dbo.Materiales.NOMBRE, dbo.Materiales.STOCKACTUAL, dbo.Materiales.PRECIOCOSTEPTA, dbo.Materiales.PRECIOCOSTEEUR, 
                      (dbo.PedidoVentaItems.CantidadT36 + dbo.PedidoVentaItems.CantidadT38 + dbo.PedidoVentaItems.CantidadT40 + dbo.PedidoVentaItems.CantidadT42 +
                       dbo.PedidoVentaItems.CantidadT44 + dbo.PedidoVentaItems.CantidadT46 + dbo.PedidoVentaItems.CantidadT48 + dbo.PedidoVentaItems.CantidadT50 +
                       dbo.PedidoVentaItems.CantidadT52 + dbo.PedidoVentaItems.CantidadT54 + dbo.PedidoVentaItems.CantidadT56) 
                      * dbo.EstrModelos.CANTIDAD AS Cantidad, 
                      (dbo.PedidoVentaItems.ServidoT36 + dbo.PedidoVentaItems.ServidoT38 + dbo.PedidoVentaItems.ServidoT40 + dbo.PedidoVentaItems.ServidoT42 + dbo.PedidoVentaItems.ServidoT44
                       + dbo.PedidoVentaItems.ServidoT46 + dbo.PedidoVentaItems.ServidoT48 + dbo.PedidoVentaItems.ServidoT50 + dbo.PedidoVentaItems.ServidoT52 + dbo.PedidoVentaItems.ServidoT54
                       + dbo.PedidoVentaItems.ServidoT56) * dbo.Modelos.CANTIDADTELA AS Servido, dbo.PedidoVentaItems.TemporadaID, dbo.Materiales.AnchuraTela, 
                      dbo.PedidosVenta.Numero, dbo.Proveedores.Nombre AS NombreProveedor, CAST('' AS VARCHAR(30)) AS NombreColor, '' AS NombreSerie, 
                      0 AS SerieID
FROM         dbo.PedidoVentaItems INNER JOIN
                      dbo.ArticuloColores ON dbo.PedidoVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID INNER JOIN
                      dbo.Articulos ON dbo.ArticuloColores.ARTICULOID = dbo.Articulos.ARTICULOID INNER JOIN
                      dbo.Modelos ON dbo.Articulos.MODELOID = dbo.Modelos.MODELOID INNER JOIN
                      dbo.EstrModelos ON dbo.Modelos.MODELOID = dbo.EstrModelos.MODELOID INNER JOIN
                      dbo.Materiales ON dbo.EstrModelos.MATERIALID = dbo.Materiales.MATERIALID INNER JOIN
                      dbo.PedidosVenta ON dbo.PedidoVentaItems.PedidoVentaID = dbo.PedidosVenta.PedidoVentaID LEFT OUTER JOIN
                      dbo.Proveedores ON dbo.Materiales.ProveedorID = dbo.Proveedores.ProveedorID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spNecesidadesMaterial]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spNecesidadesMaterial]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto:  procedimiento  almacenado dbo.spNecesidadesMaterial    fecha de la secuencia de comandos: 10/09/2001 18:06:22 ******/
CREATE PROCEDURE [spNecesidadesMaterial] 

AS

SELECT *
from vNecesidadesMaterial1
UNION
SELECT *
from vNecesidadesMaterial2
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


--- añadir precio de compra...
ALTER TABLE Articulos ADD 
	[PrecioCompraEUR] [IMPORTEEUR] NULL
GO
UPDATE Articulos SET PrecioCompraEUR = 0 
GO

ALTER TABLE Articulos ALTER COLUMN [PrecioCompraEUR] [IMPORTEEUR] NOT NULL
GO
setuser
GO

EXEC sp_bindefault N'[dbo].[0]', N'[Articulos].[PrecioCompraEUR]'
GO

setuser
GO

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
                      dbo.Articulos.SERIEID, dbo.Articulos.SuReferencia, dbo.Articulos.ProveedorID, dbo.Articulos.PrecioCompraEUR
FROM         dbo.Articulos INNER JOIN
                      dbo.Series ON dbo.Articulos.SERIEID = dbo.Series.SerieID INNER JOIN
                      dbo.Prendas ON dbo.Articulos.PRENDAID = dbo.Prendas.PrendaID INNER JOIN
                      dbo.Modelos ON dbo.Articulos.MODELOID = dbo.Modelos.MODELOID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

