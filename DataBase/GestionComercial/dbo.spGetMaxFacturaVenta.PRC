SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE dbo.spGetMaxFacturaVenta 
	(
		@TemporadaID INT,
		@EmpresaID INT,
		@Numero INT output
	)

AS
DECLARE @AuxNumero INT
DECLARE @ExisteFactura INT
BEGIN
	SELECT @ExisteFactura = count(*)
	FROM FacturasVenta
	WHERE Numero = @Numero
	AND TemporadaID = @TemporadaID
	AND EmpresaID = @EmpresaID

	print @Numero
	print @ExisteFactura
	
	WHILE  @ExisteFactura > 0 
	BEGIN
		SET @Numero = @Numero + 1

		SELECT @ExisteFactura = count(*)
		FROM FacturasVenta
		WHERE Numero = @Numero
		AND TemporadaID = @TemporadaID
		AND EmpresaID = @EmpresaID
	END
	print @Numero

END
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

