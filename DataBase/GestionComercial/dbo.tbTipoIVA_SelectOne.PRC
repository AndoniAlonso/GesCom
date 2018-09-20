SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%90beae03a23d6ac33cea163df508416f%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbTipoIVA_SelectOne
	(
	@TipoIVAID UniqueIdentifier
	)
	AS


	SELECT 
	[tbTipoIVA].[TipoIVAID], 
	[tbTipoIVA].[DescTipoIVA], 
	[tbTipoIVA].[PorcIVA]
	FROM 
		   [tbTipoIVA]
		
	WHERE @TipoIVAID = [tbTipoIVA].[TipoIVAID]


	
	RETURN (@@ERROR)
	


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

