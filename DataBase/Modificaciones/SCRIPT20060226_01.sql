--Asignar Banco del proveedor al pago.sql
--SELECT     dbo.CobrosPagos.*
UPDATE dbo.CobrosPagos 
SET dbo.CobrosPagos.BancoID = dbo.Proveedores.BancoID
FROM         dbo.CobrosPagos INNER JOIN
                      dbo.Proveedores ON dbo.CobrosPagos.PersonaID = dbo.Proveedores.ProveedorID
WHERE     (dbo.CobrosPagos.Tipo = 'P') AND (dbo.CobrosPagos.BancoID = 0)