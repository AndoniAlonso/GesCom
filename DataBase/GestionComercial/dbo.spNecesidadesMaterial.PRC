SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Objeto:  procedimiento  almacenado dbo.spNecesidadesMaterial    fecha de la secuencia de comandos: 10/09/2001 18:06:22 ******/
CREATE PROCEDURE [spNecesidadesMaterial] 

AS

SELECT *
from vNecesidadesMaterial1
UNION
SELECT *
from vNecesidadesMaterial2
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

