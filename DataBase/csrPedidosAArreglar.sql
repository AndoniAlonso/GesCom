DECLARE @RegistroActual INT 

SET NOCOUNT ON 

SET @SolicitudHasta = 155000 

SET @RegistroActual = 0 

DECLARE rsPedidos CURSOR LOCAL FOR 
SELECT *
FROM   vKKPedidosAArreglar


OPEN rsPedidos 

FETCH NEXT FROM rsPedidos INTO 
        @IDLineaSolicitud, 
        @IDSolicitud, 
        @Plantilla 
        
WHILE @@fetch_status = 0 
BEGIN   
   IF @RegistroActual = 0 
   BEGIN 
      BEGIN TRANSACTION 
      SET @ErrorAcumulado = 0 
      SET @ReturnCode = 0 
   END 
   SET @RegistroActual = @RegistroActual + 1 

   PRINT 'IDSolicitudLinea: ' + CAST(@IDLineaSolicitud as VARCHAR(10)) + ' ' + CAST(@IDSolicitud as VARCHAR(10)) + ' ' + cast(getdate() as VARCHAR(25))

   -- Traspasamos el registro de la cabecera de solicitud (siempre que no exista ya) 
   INSERT INTO tbSolicitudCabeceraHist 
    SELECT * FROM tbSolicitudCabecera 
    WHERE IDSolicitud = @IDSolicitud 
    AND NOT EXISTS 
        (SELECT * 
         FROM tbSolicitudCabeceraHist 
         WHERE IDSolicitud = @IDSolicitud) 
   SET @ReturnCode = @@ERROR 
   SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

   -- Traspasamos el registro de la linea de solicitud 
   INSERT INTO tbSolicitudLineaHist 
    SELECT * FROM tbSolicitudLinea 
    WHERE IDLineaSolicitud = @IDLineaSolicitud 
   SET @ReturnCode = @@ERROR 
   SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

   SELECT @HayExpedicion=COUNT(*) 
   FROM   tbExpedicionLinea 
   WHERE IDLineaSolicitud = @IDLineaSolicitud 

   IF @HayExpedicion > 0 
   BEGIN 
       -- Leemos el nº de cabecera de expedición 
       SELECT @IDExpedicion=IDExpedicion 
       FROM   tbExpedicionLinea 
       WHERE IDLineaSolicitud = @IDLineaSolicitud 

       -- Borramos el registro de linea de expedición traspasado. 
       DELETE FROM tbExpedicionLinea 
       WHERE IDLineaSolicitud = @IDLineaSolicitud 
       SET @ReturnCode = @@ERROR 
       SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

       -- Borramos la cabecera de expedición si ya no tiene lineas 
       DELETE FROM tbExpedicionCabecera 
       WHERE IDExpedicion = @IDExpedicion 
       AND NOT EXISTS (SELECT * 
                      FROM tbExpedicionLinea 
                      WHERE IDExpedicion = @IDExpedicion) 
       SET @ReturnCode = @@ERROR 
       SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 
   END 

   -- Borramos el registro de Pedidos agrupados. 
   DELETE FROM tbSolicitudLinPedAgrupado 
   WHERE IDLineaSolicitud = @IDLineaSolicitud 
   SET @ReturnCode = @@ERROR 
   SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

   -- Borramos el registro de ofertas traspasadas. 
   DELETE FROM tbSolicitudLinOfAgrupada 
   WHERE IDLineaSolicitud = @IDLineaSolicitud 
   SET @ReturnCode = @@ERROR 
   SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

   IF @Plantilla = 1 
   BEGIN 
      -- Traspasamos los registros de plantillas 
      INSERT INTO tbSolicitudEspecificacionHist 
       SELECT * FROM tbSolicitudEspecificacion 
       WHERE IDLinea = @IDLineaSolicitud 
      SET @ReturnCode = @@ERROR 
      SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

      -- Borramos el registro de especificacion traspasado. 
      DELETE FROM tbSolicitudEspecificacion 
      WHERE IDLinea = @IDLineaSolicitud 
      SET @ReturnCode = @@ERROR 
      SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 
   END 

   -- Borramos el registro de linea traspasado. 
   DELETE FROM tbSolicitudLinea 
   WHERE IDLineaSolicitud = @IDLineaSolicitud 
   SET @ReturnCode = @@ERROR 
   SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 

   -- Borramos el registro de cabecera de solicitud 
   DELETE FROM tbSolicitudCabecera 
   WHERE IDSolicitud = @IDSolicitud 
   AND NOT EXISTS 
       (SELECT * 
        FROM tbSolicitudLinea 
        WHERE IDSolicitud = @IDSolicitud) 
   SET @ReturnCode = @@ERROR 
   SET @ErrorAcumulado = @ErrorAcumulado + @ReturnCode 
   
   IF @ErrorAcumulado = 0 
   BEGIN 
      COMMIT TRANSACTION 
   END 
   ELSE 
   BEGIN 
      ROLLBACK TRANSACTION 
   END 
   END 



   FETCH NEXT FROM rsPedidos INTO 
        @IDLineaSolicitud, 
        @IDSolicitud, 
        @Plantilla 
END 
IF @RegistroActual <> 0 
BEGIN 
   IF @ErrorAcumulado = 0 
   BEGIN 
      COMMIT TRANSACTION 
   END 
   ELSE 
   BEGIN 
      ROLLBACK TRANSACTION 
   END 
END 

CLOSE rsPedidos 
DEALLOCATE rsPedidos 

GO 
