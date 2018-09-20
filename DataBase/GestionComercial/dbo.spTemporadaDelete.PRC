SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto:  procedimiento  almacenado dbo.spTemporadaDelete    fecha de la secuencia de comandos: 10/09/2001 18:06:22 ******/
CREATE PROCEDURE spTemporadaDelete

@TemporadaID int

AS

DELETE FROM Temporadas

WHERE TemporadaID = @TemporadaID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

