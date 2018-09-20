SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/* ***^^^***|||***^^^*** *//*  *//*  *//*  *//* %%%###%%%3a0671904e6669ffdbf3b6f6306103c8%%%###%%% *//* ***^^^***|||***^^^*** */

CREATE PROCEDURE tbOperario_Insert
	(
	
	@OperarioID UniqueIdentifier, 
	@DescOperario VarChar (50), 
	@PrecioHora Decimal, 
	@Descuento Decimal, 
	@Activo Bit
	)
	AS
   INSERT INTO  [tbOperario]
	(
	[OperarioID], 
	[DescOperario], 
	[PrecioHora], 
	[Descuento], 
	[Activo]
	)
	VALUES
	(
	@OperarioID, 
	@DescOperario, 
	@PrecioHora, 
	@Descuento, 
	@Activo
	)
	
		
	
	RETURN (@@ERROR)
	
	

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

