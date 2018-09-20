INSERT INTO ParametrosAplicacion values ('TALLPRED', 'Tallaje predeterminado estandar', 1, 1, 1)
GO
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vAlbaranesCompra]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vAlbaranesCompra]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.vAlbaranesCompra
AS
SELECT     dbo.AlbaranesCompra.AlbaranCompraID, dbo.AlbaranesCompra.Fecha, dbo.AlbaranesCompra.Numero, RTRIM(dbo.AlbaranesCompra.Observaciones) 
                      AS Observaciones, dbo.AlbaranesCompra.TotalBrutoPTA, dbo.AlbaranesCompra.TotalBrutoEUR, dbo.AlbaranesCompra.TemporadaID, 
                      dbo.AlbaranesCompra.EmpresaID, RTRIM(dbo.Transportistas.Nombre) AS NombreTransportista, RTRIM(dbo.Proveedores.Nombre) 
                      AS NombreProveedor, dbo.Proveedores.ProveedorID, 
                      SUM(dbo.AlbaranCompraArticulos.CantidadT36 + dbo.AlbaranCompraArticulos.CantidadT38 + dbo.AlbaranCompraArticulos.CantidadT40 + dbo.AlbaranCompraArticulos.CantidadT42
                       + dbo.AlbaranCompraArticulos.CantidadT44 + dbo.AlbaranCompraArticulos.CantidadT46 + dbo.AlbaranCompraArticulos.CantidadT48 + dbo.AlbaranCompraArticulos.CantidadT50
                       + dbo.AlbaranCompraArticulos.CantidadT52 + dbo.AlbaranCompraArticulos.CantidadT54 + dbo.AlbaranCompraArticulos.CantidadT56) 
                      AS CantidadArticulos
FROM         dbo.AlbaranesCompra INNER JOIN
                      dbo.Proveedores ON dbo.AlbaranesCompra.ProveedorID = dbo.Proveedores.ProveedorID LEFT OUTER JOIN
                      dbo.AlbaranCompraArticulos ON dbo.AlbaranesCompra.AlbaranCompraID = dbo.AlbaranCompraArticulos.AlbaranCompraID LEFT OUTER JOIN
                      dbo.Transportistas ON dbo.AlbaranesCompra.TransportistaID = dbo.Transportistas.TransportistaID
GROUP BY dbo.AlbaranesCompra.AlbaranCompraID, dbo.AlbaranesCompra.Fecha, dbo.AlbaranesCompra.Numero, RTRIM(dbo.AlbaranesCompra.Observaciones), 
                      dbo.AlbaranesCompra.TotalBrutoPTA, dbo.AlbaranesCompra.TotalBrutoEUR, dbo.AlbaranesCompra.TemporadaID, dbo.AlbaranesCompra.EmpresaID, 
                      RTRIM(dbo.Transportistas.Nombre), RTRIM(dbo.Proveedores.Nombre), dbo.Proveedores.ProveedorID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

