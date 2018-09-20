SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%6b8c7139e042fc3786dac7eb4465c26e%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbOperario_Delete
	(
	
	@OperarioID UniqueIdentifier
	)
	AS

   BEGIN
 	SET NOCOUNT ON 

   
      DELETE FROM [tbOperario]
      WHERE @OperarioID = [tbOperario].[OperarioID]



   END
				
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

