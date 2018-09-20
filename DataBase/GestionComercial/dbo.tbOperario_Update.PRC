SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%895a71b4e802e93e3338ea4bb8060bf4%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbOperario_Update
	(
	
	@OperarioID UniqueIdentifier, 
	@DescOperario VarChar (50), 
	@PrecioHora Decimal, 
	@Descuento Decimal, 
	@Activo Bit
	)
	AS

BEGIN
SET NOCOUNT ON 

	UPDATE [tbOperario]
		SET
			
		[OperarioID] = @OperarioID, 
		[DescOperario] = @DescOperario, 
		[PrecioHora] = @PrecioHora, 
		[Descuento] = @Descuento, 
		[Activo] = @Activo
		WHERE @OperarioID = [tbOperario].[OperarioID]
		
END
				
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

