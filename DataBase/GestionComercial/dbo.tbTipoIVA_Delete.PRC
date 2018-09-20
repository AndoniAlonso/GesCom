SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%ef60e9c8851e15d205c852daf4cde4b0%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbTipoIVA_Delete
	(
	
	@TipoIVAID UniqueIdentifier
	)
	AS

   BEGIN
 	SET NOCOUNT ON 

   
      DELETE FROM [tbTipoIVA]
      WHERE @TipoIVAID = [tbTipoIVA].[TipoIVAID]



   END
				
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

