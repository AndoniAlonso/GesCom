-- Creamos la tabla de versiones
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[eitDatabaseVersion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[eitDatabaseVersion]
GO

CREATE TABLE [dbo].[eitDatabaseVersion] (
	[DatabaseVersionID] [int] IDENTITY (1, 1) NOT NULL ,
	[Version] [int] NOT NULL ,
	[Fecha] [datetime] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[eitDatabaseVersion] WITH NOCHECK ADD 
	CONSTRAINT [PK_eitDatabaseVersion] PRIMARY KEY  CLUSTERED 
	(
		[DatabaseVersionID]
	)  ON [PRIMARY] 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[eispNuevaVersion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[eispNuevaVersion]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto:  procedimiento  almacenado dbo.eispNuevaVersion  ******/
CREATE PROCEDURE [eispNuevaVersion]
   @NuevaVersion int
AS
BEGIN
DECLARE @VersionActual int

	SELECT @VersionActual=COALESCE(MAX(Version),0)
	FROM eitDatabaseVersion

	IF @VersionActual = @NuevaVersion -1 
	BEGIN
		INSERT INTO eitDatabaseVersion
		VALUES (@NuevaVersion, GetDate())
		
	END 
	ELSE
	BEGIN
		RAISERROR ('No se puede establecer la siguiente versión de la base de datos a %u porque la versión actual es la nº %u',20,1,@NuevaVersion, @VersionActual)
		WITH LOG, NOWAIT
	END
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

--Establecemos la primera versión.
exec eispNuevaVersion 1
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbCoreConsultaDW]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbCoreConsultaDW]
GO

CREATE TABLE [dbo].[tbCoreConsultaDW] (
	[ConsultaID] [char] (10) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[Descripcion] [varchar] (100) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[TiempoRefresco] [int] NOT NULL ,
	[DatePartRefresco] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[VistaOrigen] [varchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[TablaDestino] [varchar] (50) COLLATE Modern_Spanish_CI_AS NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tbCoreConsultaDW] WITH NOCHECK ADD 
	CONSTRAINT [PK_tbCoreConsultaDW] PRIMARY KEY  CLUSTERED 
	(
		[ConsultaID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tbCoreConsultaDW] ADD 
	CONSTRAINT [DF_tbCoreConsultaDW_DatePartRefresco] DEFAULT ('dd') FOR [DatePartRefresco]
GO

INSERT INTO tbCoreConsultaDW VALUES
('CVTempProv', 'Compras y Ventas de una temporada por proveedor', 1, 'dd', 'vComprasVentasTemporadaProveedor',
 'tbDWComprasVentasTemporadaProveedor')
GO





if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vVentasTiendaProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vVentasTiendaProveedor]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vVentasTiendaProveedor
AS
SELECT     EnumeradoTipoProveedor.NombMiembro, SUM(dbo.FacturaVentaItems.Cantidad) AS Cantidad, dbo.FacturasVenta.CentroGestionID, 
                      dbo.Articulos.ProveedorID, dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre AS NombreProveedor, 
                      UnidadesCompra.Cantidad AS CantidadCompra, dbo.FacturaVentaItems.Descuento, dbo.FacturasVenta.EmpresaID, dbo.FacturasVenta.TemporadaID, 
                      dbo.ArticuloColores.ARTICULOCOLORID
FROM         (SELECT     ArticuloColorID, 
                                              SUM(CantidadT36 + CantidadT38 + CantidadT40 + CantidadT42 + CantidadT44 + CantidadT46 + CantidadT48 + CantidadT50 + CantidadT52 + CantidadT54
                                               + CantidadT56) AS Cantidad
                       FROM          dbo.PedidoCompraArticulos
                       GROUP BY ArticuloColorID) UnidadesCompra RIGHT OUTER JOIN
                      dbo.FacturaVentaItems INNER JOIN
                      dbo.FacturasVenta ON dbo.FacturaVentaItems.FacturaVentaID = dbo.FacturasVenta.FacturaVentaID ON 
                      UnidadesCompra.ArticuloColorID = dbo.FacturaVentaItems.ArticuloColorID LEFT OUTER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                          (SELECT     *
                            FROM          Enumerados
                            WHERE      NombEnum = 'enuTipoProveedor') EnumeradoTipoProveedor ON 
                      dbo.Proveedores.TipoProveedor = EnumeradoTipoProveedor.ValorConst ON 
                      dbo.FacturaVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID
GROUP BY dbo.FacturasVenta.CentroGestionID, dbo.Articulos.ProveedorID, dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre, 
                      UnidadesCompra.Cantidad, EnumeradoTipoProveedor.NombMiembro, dbo.FacturaVentaItems.Descuento, dbo.FacturasVenta.EmpresaID, 
                      dbo.FacturasVenta.TemporadaID, dbo.ArticuloColores.ARTICULOCOLORID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vComprasVentasTemporadaModeloProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vComprasVentasTemporadaModeloProveedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vComprasVentasTemporadaProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vComprasVentasTemporadaProveedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vVentasTemporadaModeloProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vVentasTemporadaModeloProveedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vVentasTemporadaProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vVentasTemporadaProveedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vVentasTemporadaProveedorTienda]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vVentasTemporadaProveedorTienda]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vComprasTemporadaModeloProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vComprasTemporadaModeloProveedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vComprasTemporadaProveedor]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vComprasTemporadaProveedor]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vComprasTemporadaModeloProveedor
AS
SELECT     TOP 100 PERCENT dbo.Proveedores.Nombre, dbo.Prendas.NOMBRE AS NombrePrenda, dbo.Modelos.NOMBRE AS NombreModelo, 
                      dbo.Series.Nombre AS NombreSerie, 
                      SUM(dbo.AlbaranCompraArticulos.CantidadT36 + dbo.AlbaranCompraArticulos.CantidadT38 + dbo.AlbaranCompraArticulos.CantidadT40 + dbo.AlbaranCompraArticulos.CantidadT42
                       + dbo.AlbaranCompraArticulos.CantidadT44 + dbo.AlbaranCompraArticulos.CantidadT46 + dbo.AlbaranCompraArticulos.CantidadT48 + dbo.AlbaranCompraArticulos.CantidadT50
                       + dbo.AlbaranCompraArticulos.CantidadT52 + dbo.AlbaranCompraArticulos.CantidadT54 + dbo.AlbaranCompraArticulos.CantidadT56) AS Cantidad, 
                      dbo.Enumerados.NombMiembro, dbo.Temporadas.Nombre AS NombreTemporada, dbo.Proveedores.ProveedorID
FROM         dbo.Temporadas INNER JOIN
                      dbo.AlbaranCompraArticulos ON dbo.Temporadas.TemporadaID = dbo.AlbaranCompraArticulos.TemporadaID INNER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                      dbo.Enumerados ON dbo.Proveedores.TipoProveedor = dbo.Enumerados.ValorConst ON 
                      dbo.AlbaranCompraArticulos.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID INNER JOIN
                      dbo.Prendas ON dbo.Articulos.PRENDAID = dbo.Prendas.PrendaID INNER JOIN
                      dbo.Modelos ON dbo.Articulos.MODELOID = dbo.Modelos.MODELOID INNER JOIN
                      dbo.Series ON dbo.Articulos.SERIEID = dbo.Series.SerieID
WHERE     (dbo.AlbaranCompraArticulos.TemporadaID = 2)
GROUP BY dbo.Proveedores.Nombre, dbo.Proveedores.Codigo, dbo.Proveedores.TipoProveedor, dbo.Enumerados.NombMiembro, dbo.Enumerados.NombEnum,
                       dbo.Temporadas.Nombre, dbo.Proveedores.ProveedorID, dbo.Prendas.NOMBRE, dbo.Series.Nombre, dbo.Modelos.NOMBRE
HAVING      (dbo.Enumerados.NombEnum = 'enuTipoProveedor')
ORDER BY dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vComprasTemporadaProveedor
AS
SELECT     TOP 100 PERCENT dbo.Proveedores.Nombre, 
                      SUM(dbo.AlbaranCompraArticulos.CantidadT36 + dbo.AlbaranCompraArticulos.CantidadT38 + dbo.AlbaranCompraArticulos.CantidadT40 + dbo.AlbaranCompraArticulos.CantidadT42
                       + dbo.AlbaranCompraArticulos.CantidadT44 + dbo.AlbaranCompraArticulos.CantidadT46 + dbo.AlbaranCompraArticulos.CantidadT48 + dbo.AlbaranCompraArticulos.CantidadT50
                       + dbo.AlbaranCompraArticulos.CantidadT52 + dbo.AlbaranCompraArticulos.CantidadT54 + dbo.AlbaranCompraArticulos.CantidadT56) AS Cantidad, 
                      dbo.Enumerados.NombMiembro, dbo.Temporadas.Nombre AS NombreTemporada, dbo.Proveedores.ProveedorID
FROM         dbo.Temporadas INNER JOIN
                      dbo.AlbaranCompraArticulos ON dbo.Temporadas.TemporadaID = dbo.AlbaranCompraArticulos.TemporadaID INNER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                      dbo.Enumerados ON dbo.Proveedores.TipoProveedor = dbo.Enumerados.ValorConst ON 
                      dbo.AlbaranCompraArticulos.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID
WHERE     (dbo.AlbaranCompraArticulos.TemporadaID = 2)
GROUP BY dbo.Proveedores.Nombre, dbo.Proveedores.Codigo, dbo.Proveedores.TipoProveedor, dbo.Enumerados.NombMiembro, dbo.Enumerados.NombEnum,
                       dbo.Temporadas.Nombre, dbo.Proveedores.ProveedorID
HAVING      (dbo.Enumerados.NombEnum = 'enuTipoProveedor')
ORDER BY dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vVentasTemporadaModeloProveedor
AS
SELECT     TOP 100 PERCENT dbo.Proveedores.Nombre, dbo.Prendas.NOMBRE AS NombrePrenda, dbo.Modelos.NOMBRE AS NombreModelo, 
                      dbo.Series.Nombre AS NombreSerie, SUM(dbo.FacturaVentaItems.Cantidad) AS Cantidad, dbo.Enumerados.NombMiembro, 
                      dbo.Temporadas.Nombre AS NombreTemporada, dbo.Proveedores.ProveedorID
FROM         dbo.Temporadas INNER JOIN
                      dbo.FacturaVentaItems ON dbo.Temporadas.TemporadaID = dbo.FacturaVentaItems.TemporadaID INNER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                      dbo.Enumerados ON dbo.Proveedores.TipoProveedor = dbo.Enumerados.ValorConst ON 
                      dbo.FacturaVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID INNER JOIN
                      dbo.Modelos ON dbo.Articulos.MODELOID = dbo.Modelos.MODELOID INNER JOIN
                      dbo.Series ON dbo.Articulos.SERIEID = dbo.Series.SerieID INNER JOIN
                      dbo.Prendas ON dbo.Articulos.PRENDAID = dbo.Prendas.PrendaID
WHERE     (dbo.FacturaVentaItems.TemporadaID = 2)
GROUP BY dbo.Proveedores.Nombre, dbo.Proveedores.Codigo, dbo.Proveedores.TipoProveedor, dbo.Enumerados.NombMiembro, dbo.Enumerados.NombEnum,
                       dbo.Temporadas.Nombre, dbo.Proveedores.ProveedorID, dbo.Modelos.NOMBRE, dbo.Series.Nombre, dbo.Prendas.NOMBRE
HAVING      (dbo.Enumerados.NombEnum = 'enuTipoProveedor')
ORDER BY dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre, dbo.Prendas.NOMBRE

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vComprasVentasTemporadaModeloProveedor
AS
SELECT     dbo.vComprasTemporadaModeloProveedor.Nombre as NombreProveedor, dbo.vComprasTemporadaModeloProveedor.NombrePrenda, 
                      dbo.vComprasTemporadaModeloProveedor.NombreModelo, dbo.vComprasTemporadaModeloProveedor.NombreSerie, 
                      dbo.vComprasTemporadaModeloProveedor.Cantidad AS Compras, dbo.vVentasTemporadaModeloProveedor.Cantidad AS Ventas, 
                      dbo.vComprasTemporadaModeloProveedor.NombMiembro as TipoProveedor, dbo.vComprasTemporadaModeloProveedor.NombreTemporada
FROM         dbo.vComprasTemporadaModeloProveedor INNER JOIN
                      dbo.vVentasTemporadaModeloProveedor ON 
                      dbo.vComprasTemporadaModeloProveedor.ProveedorID = dbo.vVentasTemporadaModeloProveedor.ProveedorID AND 
                      dbo.vComprasTemporadaModeloProveedor.NombrePrenda = dbo.vVentasTemporadaModeloProveedor.NombrePrenda AND 
                      dbo.vComprasTemporadaModeloProveedor.NombreModelo = dbo.vVentasTemporadaModeloProveedor.NombreModelo AND 
                      dbo.vComprasTemporadaModeloProveedor.NombreSerie = dbo.vVentasTemporadaModeloProveedor.NombreSerie

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vVentasTemporadaProveedor
AS
SELECT     TOP 100 PERCENT dbo.Proveedores.Nombre, SUM(dbo.FacturaVentaItems.Cantidad) AS Cantidad, dbo.Enumerados.NombMiembro, 
                      dbo.Temporadas.Nombre AS NombreTemporada, dbo.Proveedores.ProveedorID
FROM         dbo.Temporadas INNER JOIN
                      dbo.FacturaVentaItems ON dbo.Temporadas.TemporadaID = dbo.FacturaVentaItems.TemporadaID INNER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                      dbo.Enumerados ON dbo.Proveedores.TipoProveedor = dbo.Enumerados.ValorConst ON 
                      dbo.FacturaVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID
WHERE     (dbo.FacturaVentaItems.TemporadaID = 2)
GROUP BY dbo.Proveedores.Nombre, dbo.Proveedores.Codigo, dbo.Proveedores.TipoProveedor, dbo.Enumerados.NombMiembro, dbo.Enumerados.NombEnum,
                       dbo.Temporadas.Nombre, dbo.Proveedores.ProveedorID
HAVING      (dbo.Enumerados.NombEnum = 'enuTipoProveedor')
ORDER BY dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vComprasVentasTemporadaProveedor
AS
SELECT     dbo.vComprasTemporadaProveedor.Nombre AS NombreProveedor, dbo.vComprasTemporadaProveedor.Cantidad AS Compras, 
                      dbo.vVentasTemporadaProveedor.Cantidad AS Ventas, dbo.vComprasTemporadaProveedor.NombMiembro as TipoProveedor, 
                      dbo.vComprasTemporadaProveedor.NombreTemporada
FROM         dbo.vComprasTemporadaProveedor INNER JOIN
                      dbo.vVentasTemporadaProveedor ON dbo.vComprasTemporadaProveedor.ProveedorID = dbo.vVentasTemporadaProveedor.ProveedorID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vVentasTemporadaProveedorTienda
AS
SELECT     TOP 100 PERCENT dbo.CentrosGestion.Nombre AS NombreTienda, dbo.Proveedores.Nombre, SUM(dbo.FacturaVentaItems.Cantidad) AS Cantidad, 
                      dbo.Enumerados.NombMiembro, dbo.Temporadas.Nombre AS NombreTemporada, dbo.Proveedores.ProveedorID
FROM         dbo.Temporadas INNER JOIN
                      dbo.FacturaVentaItems ON dbo.Temporadas.TemporadaID = dbo.FacturaVentaItems.TemporadaID INNER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                      dbo.Enumerados ON dbo.Proveedores.TipoProveedor = dbo.Enumerados.ValorConst ON 
                      dbo.FacturaVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID INNER JOIN
                      dbo.FacturasVenta ON dbo.FacturaVentaItems.FacturaVentaID = dbo.FacturasVenta.FacturaVentaID INNER JOIN
                      dbo.CentrosGestion ON dbo.FacturasVenta.CentroGestionID = dbo.CentrosGestion.CentroGestionID
WHERE     (dbo.FacturaVentaItems.TemporadaID = 2)
GROUP BY dbo.Proveedores.Nombre, dbo.Proveedores.Codigo, dbo.Proveedores.TipoProveedor, dbo.Enumerados.NombMiembro, dbo.Enumerados.NombEnum,
                       dbo.Temporadas.Nombre, dbo.Proveedores.ProveedorID, dbo.CentrosGestion.Nombre
HAVING      (dbo.Enumerados.NombEnum = 'enuTipoProveedor')
ORDER BY dbo.CentrosGestion.Nombre, dbo.Proveedores.TipoProveedor, dbo.Proveedores.Nombre

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tbDWComprasVentasTemporadaProveedor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tbDWComprasVentasTemporadaProveedor]
GO

CREATE TABLE [dbo].[tbDWComprasVentasTemporadaProveedor] (
	[NombreProveedor] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Compras] [CANTIDAD] NULL ,
	[Ventas] [CANTIDAD] NULL ,
	[TipoProveedor] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[NombreTemporada] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

