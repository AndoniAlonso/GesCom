USE GESTIONCOMERCIAL
GO 
CREATE DEFAULT Now AS GETDATE()
GO	

ALTER TABLE [dbo].[FacturasVenta] 
	ADD FechaHora DATETIME NULL,
	[EntregadoEUR] [IMPORTEEUR] NULL ,
	[CambiosEUR] [IMPORTEEUR] NULL
	
GO

UPDATE [dbo].[FacturasVenta] 
	SET FechaHora = Fecha,
		EntregadoEUR = 0,
		CambiosEUR = 0
	
GO

ALTER TABLE [dbo].[FacturasVenta] 
	ALTER COLUMN FechaHora DATETIME NOT NULL
GO
ALTER TABLE [dbo].[FacturasVenta] 
	ALTER COLUMN EntregadoEUR [IMPORTEEUR] NOT NULL
GO
ALTER TABLE [dbo].[FacturasVenta] 
	ALTER COLUMN CambiosEUR [IMPORTEEUR] NOT NULL
GO

setuser
GO

EXEC sp_bindefault N'[dbo].[0]', N'[FacturasVenta].[EntregadoEUR]'
GO

EXEC sp_bindefault N'[dbo].[0]', N'[FacturasVenta].[CambiosEUR]'
GO

EXEC sp_bindefault N'[dbo].[Now]', N'[FacturasVenta].[FechaHora]'
GO

setuser
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTPVCierreCaja]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTPVCierreCaja]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTPVCierreCaja
AS
SELECT     TOP 100 PERCENT dbo.FacturasVenta.FechaHora AS [Fecha-Hora], dbo.FacturasVenta.Numero, dbo.Representantes.Nombre AS NombreRepresentante,
                       CASE WHEN dbo.FacturasVenta.NetoEUR >= 0 THEN dbo.FacturasVenta.NetoEUR ELSE 0 END AS Importe, 
                      CASE WHEN dbo.FacturasVenta.NetoEUR < 0 THEN - 1 * dbo.FacturasVenta.NetoEUR ELSE 0 END AS Devolucion, 
                      dbo.FacturasVenta.EntregadoEUR AS Entregado, dbo.FacturasVenta.CambiosEUR AS Cambio, dbo.CentrosGestion.Nombre AS NombreCentroGestion, 
                      dbo.FacturasVenta.CentroGestionID
FROM         dbo.FacturasVenta INNER JOIN
                      dbo.Representantes ON dbo.FacturasVenta.RepresentanteID = dbo.Representantes.RepresentanteID INNER JOIN
                      dbo.CentrosGestion ON dbo.FacturasVenta.CentroGestionID = dbo.CentrosGestion.CentroGestionID
WHERE     FechaHora > GETDATE() - 365
ORDER BY FechaHora DESC

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vTPVReportVentas]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vTPVReportVentas]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vTPVReportVentas
AS
SELECT     TOP 100 PERCENT dbo.FacturasVenta.FechaHora as [Fecha-Hora], dbo.FacturasVenta.Numero, dbo.Representantes.Nombre AS [Nombre vendedor], 
                      dbo.FacturaVentaItems.ArticuloColorID as [Código artículo], dbo.FacturaVentaItems.Descripcion, dbo.FacturaVentaItems.BrutoEUR as Importe, dbo.FacturasVenta.CentroGestionID, 
                      dbo.CentrosGestion.Nombre AS [Nombre Centro]
FROM         dbo.FacturaVentaItems INNER JOIN
                      dbo.FacturasVenta ON dbo.FacturaVentaItems.FacturaVentaID = dbo.FacturasVenta.FacturaVentaID INNER JOIN
                      dbo.Representantes ON dbo.FacturasVenta.RepresentanteID = dbo.Representantes.RepresentanteID INNER JOIN
                      dbo.CentrosGestion ON dbo.FacturasVenta.CentroGestionID = dbo.CentrosGestion.CentroGestionID
WHERE FechaHora > GETDATE() - 365 
ORDER BY FechaHora DESC

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




--OJOOJO ¿HAY QUE ACTUALIZAR ESTO ASI????
UPDATE dbo.FacturasVenta
SET CentroGestionID = 1