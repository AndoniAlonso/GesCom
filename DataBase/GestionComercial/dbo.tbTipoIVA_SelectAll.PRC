SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%08fde2bb43807cfdfcff0e8ac3b16e74%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbTipoIVA_SelectAll
AS

   SELECT 
	[tbTipoIVA].[TipoIVAID], 
	[tbTipoIVA].[DescTipoIVA], 
	[tbTipoIVA].[PorcIVA]
   FROM tbTipoIVA
	
   RETURN 
	
   

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

