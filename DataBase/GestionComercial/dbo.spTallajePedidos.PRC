SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto:  procedimiento  almacenado dbo.spTallajePedido    fecha de la secuencia de comandos: 10/09/2001 18:06:22 ******/
CREATE PROCEDURE [spTallajePedidos] AS

SELECT CODIGOMODELO, NOMBREMODELO,
	CASE 
		WHEN SUM(CantidadT36) !=0	THEN 'Talla 36'
		WHEN SUM(CantidadT38) !=0	THEN 'Talla 38'
		WHEN SUM(CantidadT40) !=0	THEN 'Talla 40'
		WHEN SUM(CantidadT42) !=0	THEN 'Talla 42'
		WHEN SUM(CantidadT44) !=0	THEN 'Talla 44'
		WHEN SUM(CantidadT46) !=0	THEN 'Talla 46'
		WHEN SUM(CantidadT48) !=0	THEN 'Talla 48'
		WHEN SUM(CantidadT50) !=0	THEN 'Talla 50'
		WHEN SUM(CantidadT52) !=0	THEN 'Talla 52'
		WHEN SUM(CantidadT54) !=0	THEN 'Talla 54'
		WHEN SUM(CantidadT56) !=0	THEN 'Talla 56'
		ELSE 		'No hay tallas de este modelo'
	END as TallaMinima,
	CASE 
		WHEN SUM(CantidadT56) !=0	THEN 'Talla 56'
		WHEN SUM(CantidadT54) !=0	THEN 'Talla 54'
		WHEN SUM(CantidadT52) !=0	THEN 'Talla 52'
		WHEN SUM(CantidadT50) !=0	THEN 'Talla 50'
		WHEN SUM(CantidadT48) !=0	THEN 'Talla 48'
		WHEN SUM(CantidadT46) !=0	THEN 'Talla 46'
		WHEN SUM(CantidadT44) !=0	THEN 'Talla 44'
		WHEN SUM(CantidadT42) !=0	THEN 'Talla 42'
		WHEN SUM(CantidadT40) !=0	THEN 'Talla 40'
		WHEN SUM(CantidadT38) !=0	THEN 'Talla 38'
		WHEN SUM(CantidadT36) !=0	THEN 'Talla 36'
		ELSE 		'No hay tallas de este modelo'
	END as TallaMaxima
FROM vTallajePedidos
GROUP BY CODIGOMODELO, NOMBREMODELO
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

