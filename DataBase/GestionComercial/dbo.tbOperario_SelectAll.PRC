SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%ac217fb328cec3f4d7bdcd959a5eef26%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbOperario_SelectAll
AS

   SELECT 
	[tbOperario].[OperarioID], 
	[tbOperario].[DescOperario], 
	[tbOperario].[PrecioHora], 
	[tbOperario].[Descuento], 
	[tbOperario].[Activo]
   FROM tbOperario
	
   RETURN 
	
   

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

