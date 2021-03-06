if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vProveedores]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vProveedores]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vProveedores
AS
SELECT     dbo.Proveedores.ProveedorID, RTRIM(dbo.Proveedores.Nombre) AS Nombre, RTRIM(dbo.Proveedores.Titular) AS Titular, 
                      RTRIM(dbo.Proveedores.Contacto) AS Contacto, dbo.Proveedores.DNINIF, dbo.Proveedores.CuentaContable, RTRIM(dbo.Direcciones.Calle) AS Calle, 
                      RTRIM(dbo.Direcciones.Poblacion) AS Poblacion, RTRIM(dbo.Direcciones.Provincia) AS Provincia, RTRIM(dbo.Direcciones.Telefono1) AS Telefono1, 
                      dbo.MediosPago.NombreAbreviado, RTRIM(dbo.FormasDePago.Nombre) AS DescFormaPago, RTRIM(dbo.CuentasBancarias.NombreEntidad) 
                      AS NombreEntidad, RTRIM(dbo.Transportistas.Nombre) AS NombreTransportista
FROM         dbo.CuentasBancarias INNER JOIN
                      dbo.Bancos ON dbo.CuentasBancarias.CuentaBancariaID = dbo.Bancos.CuentaBancariaID RIGHT OUTER JOIN
                      dbo.Proveedores INNER JOIN
                      dbo.Direcciones ON dbo.Proveedores.DireccionID = dbo.Direcciones.DireccionID INNER JOIN
                      dbo.FormasDePago ON dbo.Proveedores.FormaPagoID = dbo.FormasDePago.FormaPagoID ON 
                      dbo.Bancos.BancoID = dbo.Proveedores.BancoID LEFT OUTER JOIN
                      dbo.Transportistas ON dbo.Proveedores.TransportistaID = dbo.Transportistas.TransportistaID LEFT OUTER JOIN
                      dbo.MediosPago ON dbo.Proveedores.MedioPagoID = dbo.MediosPago.MedioPagoID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

ALTER TABLE FacturasCompra 
ALTER COLUMN BancoID Integer NULL


if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[vFacturasCompra]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[vFacturasCompra]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.vFacturasCompra
AS
SELECT     dbo.FacturasCompra.FacturaCompraID, dbo.FacturasCompra.Fecha, dbo.FacturasCompra.Numero, RTRIM(dbo.FacturasCompra.Observaciones) 
                      AS Observaciones, dbo.FacturasCompra.BrutoEUR, dbo.FacturasCompra.TemporadaID, dbo.FacturasCompra.EmpresaID, 
                      RTRIM(dbo.Proveedores.Nombre) AS NombreProveedor, dbo.vBancos.NombreEntidad AS NombreBanco, RTRIM(dbo.Transportistas.Nombre) 
                      AS NombreTransportista, dbo.FacturasCompra.SituacionContable, dbo.Proveedores.CuentaContrapartida AS CuentaContable, 
                      dbo.FacturasCompra.FechaContable, dbo.FacturasCompra.IVAEUR, dbo.FacturasCompra.NetoEUR, dbo.FacturasCompra.DescuentoEUR, 
                      dbo.FacturasCompra.PortesEUR, dbo.FacturasCompra.EmbalajesEUR, STR(dbo.FacturasCompra.Numero, 7, 0) 
                      + dbo.FacturasCompra.Sufijo AS CodigoFactura, dbo.FacturasCompra.EmbalajesEUR + dbo.FacturasCompra.PortesEUR AS GastosEUR, 
                      dbo.Proveedores.ProveedorID, YEAR(dbo.FacturasCompra.FechaContable) AS Anio, dbo.FacturasCompra.BaseImponibleEUR
FROM         dbo.FacturasCompra INNER JOIN
                      dbo.Proveedores ON dbo.FacturasCompra.ProveedorID = dbo.Proveedores.ProveedorID LEFT OUTER JOIN
                      dbo.vBancos ON dbo.FacturasCompra.BancoID = dbo.vBancos.BancoID LEFT OUTER JOIN
                      dbo.Transportistas ON dbo.FacturasCompra.TransportistaID = dbo.Transportistas.TransportistaID

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

