SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto:  procedimiento  almacenado dbo.spTemporadaGet    fecha de la secuencia de comandos: 10/09/2001 18:06:22 ******/
CREATE PROCEDURE spTemporadaGet

@TemporadaID int

AS

SELECT TemporadaID,
               Nombre,
               Codigo

FROM Temporadas

WHERE TemporadaID = @TemporadaID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

