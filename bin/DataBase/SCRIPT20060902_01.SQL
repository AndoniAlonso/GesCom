USE GESTIONCOMERCIAL
GO 
ALTER TABLE [dbo].[CentrosGestion] 
	ADD EmpresaID INTEGER NULL
	
GO

UPDATE [dbo].[CentrosGestion] 
	SET EmpresaID = 2
GO

ALTER TABLE [dbo].[CentrosGestion] 
	ALTER COLUMN EmpresaID INTEGER NOT NULL
GO

ALTER TABLE [dbo].[CentrosGestion] 
	ADD 
	CONSTRAINT [FK_CentrosGestion_Empresas] FOREIGN KEY 
	(
		[EmpresaID]
	) REFERENCES [dbo].[Empresas] (
		[EmpresaID]
	)
GO



ALTER TABLE [dbo].[AlbaranesVenta] 
	ADD MedioPagoID INTEGER NULL
GO

UPDATE [dbo].[AlbaranesVenta] 
	SET MedioPagoID = 1
GO

ALTER TABLE [dbo].[AlbaranesVenta] 
	ALTER COLUMN MedioPagoID INTEGER NOT NULL
GO

ALTER TABLE [dbo].[AlbaranesVenta] 
	ADD 
	CONSTRAINT [FK_AlbaranesVenta_MediosPago] FOREIGN KEY 
	(
		[MedioPagoID]
	) REFERENCES [dbo].[MediosPago] (
		[MedioPagoID]
	)
GO


ALTER TABLE [dbo].[FacturasVenta] 
	ADD MedioPagoID INTEGER NULL
GO

UPDATE [dbo].[FacturasVenta] 
	SET MedioPagoID = 1
GO

ALTER TABLE [dbo].[FacturasVenta] 
	ALTER COLUMN MedioPagoID INTEGER NOT NULL
GO

ALTER TABLE [dbo].[FacturasVenta] 
	ADD 
	CONSTRAINT [FK_FacturasVenta_MediosPago] FOREIGN KEY 
	(
		[MedioPagoID]
	) REFERENCES [dbo].[MediosPago] (
		[MedioPagoID]
	)
GO

-- Maestro de proveedores
ALTER TABLE [dbo].[Proveedores] 
	ADD Codigo CHAR(3) NULL
GO

ALTER TABLE [dbo].[Proveedores] 
	ADD TipoProveedor CHAR(1) NULL
GO

UPDATE Proveedores
SET Codigo = ' ',
    TipoProveedor = ' '
GO

ALTER TABLE [dbo].[Proveedores] 
	ALTER COLUMN Codigo CHAR(3) NOT NULL
GO

ALTER TABLE [dbo].[Proveedores] 
	ALTER COLUMN TipoProveedor CHAR(1) NOT NULL
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vEstadisticaFacturaVenta]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vEstadisticaFacturaVenta]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vEstadisticaFacturaVenta
AS
SELECT     dbo.FacturaVentaItems.FacturaVentaItemID AS EstadisticaVentaID, dbo.FacturaVentaItems.Cantidad, dbo.FacturaVentaItems.BrutoEUR, 
                      dbo.FacturasVenta.Fecha, dbo.Articulos.ARTICULOID, RTRIM(dbo.Articulos.NOMBRE) AS NombreArticulo, dbo.ArticuloColores.ARTICULOCOLORID, 
                      RTRIM(dbo.ArticuloColores.NOMBRE) AS NombreArticuloColor, dbo.Modelos.MODELOID, RTRIM(dbo.Modelos.NOMBRE) AS NombreModelo, 
                      dbo.Series.SerieID, RTRIM(dbo.Series.Nombre) AS NombreSerie, dbo.Prendas.PrendaID, RTRIM(dbo.Prendas.NOMBRE) AS NombrePrenda, 
                      dbo.Representantes.RepresentanteID, RTRIM(dbo.Representantes.Nombre) AS NombreRepresentante, dbo.Clientes.ClienteID, 
                      RTRIM(dbo.Clientes.Nombre) AS NombreCliente, dbo.Temporadas.TemporadaID, RTRIM(dbo.Temporadas.Nombre) AS NombreTemporada, 
                      dbo.Empresas.EmpresaID, dbo.Empresas.Nombre AS NombreEmpresa, dbo.Direcciones.Provincia AS NombreProvincia, 
                      dbo.Articulos.PRECIOCOSTEEUR, dbo.FacturasVenta.CentroGestionID, dbo.CentrosGestion.Nombre AS NombreCentroGestion, 
                      dbo.Articulos.ProveedorID, dbo.Proveedores.Nombre AS NombreProveedor
FROM         dbo.FacturaVentaItems INNER JOIN
                      dbo.FacturasVenta ON dbo.FacturaVentaItems.FacturaVentaID = dbo.FacturasVenta.FacturaVentaID INNER JOIN
                      dbo.Empresas ON dbo.FacturasVenta.EmpresaID = dbo.Empresas.EmpresaID INNER JOIN
                      dbo.Temporadas ON dbo.FacturasVenta.TemporadaID = dbo.Temporadas.TemporadaID INNER JOIN
                      dbo.ArticuloColores ON dbo.FacturaVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID INNER JOIN
                      dbo.Articulos ON dbo.ArticuloColores.ARTICULOID = dbo.Articulos.ARTICULOID INNER JOIN
                      dbo.Modelos ON dbo.Temporadas.TemporadaID = dbo.Modelos.TEMPORADAID AND dbo.Articulos.MODELOID = dbo.Modelos.MODELOID INNER JOIN
                      dbo.Series ON dbo.Articulos.SERIEID = dbo.Series.SerieID INNER JOIN
                      dbo.Prendas ON dbo.Articulos.PRENDAID = dbo.Prendas.PrendaID INNER JOIN
                      dbo.Representantes ON dbo.FacturasVenta.RepresentanteID = dbo.Representantes.RepresentanteID INNER JOIN
                      dbo.Clientes ON dbo.FacturasVenta.ClienteID = dbo.Clientes.ClienteID INNER JOIN
                      dbo.Direcciones ON dbo.Clientes.DireccionFiscalID = dbo.Direcciones.DireccionID LEFT OUTER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID LEFT OUTER JOIN
                      dbo.CentrosGestion ON dbo.FacturasVenta.CentroGestionID = dbo.CentrosGestion.CentroGestionID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

select * 
from Articulos 
where proveedorID is null
