SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%887b7cef4bc0f29b062ea85754ed8502%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbTipoIVA_Update
	(
	
	@TipoIVAID UniqueIdentifier, 
	@DescTipoIVA VarChar (50), 
	@PorcIVA Decimal
	)
	AS

BEGIN
SET NOCOUNT ON 

	UPDATE [tbTipoIVA]
		SET
			
		[TipoIVAID] = @TipoIVAID, 
		[DescTipoIVA] = @DescTipoIVA, 
		[PorcIVA] = @PorcIVA
		WHERE @TipoIVAID = [tbTipoIVA].[TipoIVAID]
		
END
				
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

