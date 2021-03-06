if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vPedidosCompra]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vPedidosCompra]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vPedidosCompra
AS
SELECT     dbo.PedidosCompra.PedidoCompraID, dbo.PedidosCompra.Fecha, dbo.PedidosCompra.FechaEntrega, dbo.PedidosCompra.Numero, 
                      dbo.PedidosCompra.Observaciones, dbo.vBancos.NombreEmpresa AS NombreBanco, dbo.Proveedores.Nombre AS NombreProveedor, 
                      dbo.Transportistas.Nombre AS NombreTransportista, dbo.PedidosCompra.TemporadaID, dbo.PedidosCompra.EmpresaID, 
                      dbo.PedidosCompra.TotalBrutoPTA, dbo.PedidosCompra.TotalBrutoEUR
FROM         dbo.PedidosCompra INNER JOIN
                      dbo.Proveedores ON dbo.PedidosCompra.ProveedorID = dbo.Proveedores.ProveedorID LEFT OUTER JOIN
                      dbo.vBancos ON dbo.PedidosCompra.BancoID = dbo.vBancos.BancoID LEFT OUTER JOIN
                      dbo.Transportistas ON dbo.PedidosCompra.TransportistaID = dbo.Transportistas.TransportistaID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

