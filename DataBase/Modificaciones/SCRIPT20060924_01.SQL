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
                      AS NombreEntidad, RTRIM(dbo.Transportistas.Nombre) AS NombreTransportista, dbo.Proveedores.Codigo
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

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Enumerados]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Enumerados]
GO

CREATE TABLE [dbo].[Enumerados] (
	[NombEnum] [varchar] (100) NOT NULL ,
	[NombMiembro] [varchar] (100) NOT NULL ,
	[ValorConst] [int] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Enumerados] WITH NOCHECK ADD 
	CONSTRAINT [PK_Enumerados] PRIMARY KEY  CLUSTERED 
	(
		[NombEnum],
		[NombMiembro]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

INSERT INTO Enumerados VALUES ('enuTipoProveedor','Hongo',1)
INSERT INTO Enumerados VALUES ('enuTipoProveedor','Muestrario',2)
INSERT INTO Enumerados VALUES ('enuTipoProveedor','Paris-Madrid',3)
GO
UPDATE Proveedores
SET TipoProveedor = '2'
go

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
                      dbo.CentrosGestion.Nombre AS NombreCentroGestion, dbo.Articulos.ProveedorID, dbo.Proveedores.TipoProveedor, 
                      dbo.Proveedores.Nombre AS NombreProveedor, UnidadesCompra.Cantidad AS CantidadCompra, dbo.FacturaVentaItems.Descuento, 
                      dbo.FacturasVenta.EmpresaID, dbo.FacturasVenta.TemporadaID, dbo.FacturasVenta.RepresentanteID, 
                      dbo.Representantes.Nombre AS NombreRepresentante
FROM         dbo.CentrosGestion RIGHT OUTER JOIN
                          (SELECT     ArticuloColorID, 
                                                   SUM(CantidadT36 + CantidadT38 + CantidadT40 + CantidadT42 + CantidadT44 + CantidadT46 + CantidadT48 + CantidadT50 + CantidadT52 +
                                                    CantidadT54 + CantidadT56) AS Cantidad
                            FROM          dbo.PedidoCompraArticulos
                            GROUP BY ArticuloColorID) UnidadesCompra RIGHT OUTER JOIN
                      dbo.FacturaVentaItems INNER JOIN
                      dbo.FacturasVenta ON dbo.FacturaVentaItems.FacturaVentaID = dbo.FacturasVenta.FacturaVentaID INNER JOIN
                      dbo.Representantes ON dbo.FacturasVenta.RepresentanteID = dbo.Representantes.RepresentanteID ON 
                      UnidadesCompra.ArticuloColorID = dbo.FacturaVentaItems.ArticuloColorID ON 
                      dbo.CentrosGestion.CentroGestionID = dbo.FacturasVenta.CentroGestionID LEFT OUTER JOIN
                      dbo.Articulos INNER JOIN
                      dbo.ArticuloColores ON dbo.Articulos.ARTICULOID = dbo.ArticuloColores.ARTICULOID INNER JOIN
                      dbo.Proveedores ON dbo.Articulos.ProveedorID = dbo.Proveedores.ProveedorID INNER JOIN
                          (SELECT     *
                            FROM          Enumerados
                            WHERE      NombEnum = 'enuTipoProveedor') EnumeradoTipoProveedor ON 
                      dbo.Proveedores.TipoProveedor = EnumeradoTipoProveedor.ValorConst ON 
                      dbo.FacturaVentaItems.ArticuloColorID = dbo.ArticuloColores.ARTICULOCOLORID
GROUP BY dbo.FacturasVenta.CentroGestionID, dbo.CentrosGestion.Nombre, dbo.Articulos.ProveedorID, dbo.Proveedores.TipoProveedor, 
                      dbo.Proveedores.Nombre, UnidadesCompra.Cantidad, EnumeradoTipoProveedor.NombMiembro, dbo.FacturaVentaItems.Descuento, 
                      dbo.FacturasVenta.EmpresaID, dbo.FacturasVenta.TemporadaID, dbo.FacturasVenta.RepresentanteID, dbo.Representantes.Nombre

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

