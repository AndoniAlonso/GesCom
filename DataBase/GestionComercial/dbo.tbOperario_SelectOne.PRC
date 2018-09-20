SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%778dd2d0dbf3ebca6cf47ad2ceb1deb7%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbOperario_SelectOne
	(
	@OperarioID UniqueIdentifier
	)
	AS


	SELECT 
	[tbOperario].[OperarioID], 
	[tbOperario].[DescOperario], 
	[tbOperario].[PrecioHora], 
	[tbOperario].[Descuento], 
	[tbOperario].[Activo]
	FROM 
		   [tbOperario]
		
	WHERE @OperarioID = [tbOperario].[OperarioID]


	
	RETURN (@@ERROR)
	


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

