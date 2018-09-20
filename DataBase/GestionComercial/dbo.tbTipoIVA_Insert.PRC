SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%29311c5702acd9c0c6d8ce39f1ff839b%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbTipoIVA_Insert
	(
	
	@TipoIVAID UniqueIdentifier, 
	@DescTipoIVA VarChar (50), 
	@PorcIVA Decimal
	)
	AS
   INSERT INTO  [tbTipoIVA]
	(
	[TipoIVAID], 
	[DescTipoIVA], 
	[PorcIVA]
	)
	VALUES
	(
	@TipoIVAID, 
	@DescTipoIVA, 
	@PorcIVA
	)
	
		
	
	RETURN (@@ERROR)
	
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

