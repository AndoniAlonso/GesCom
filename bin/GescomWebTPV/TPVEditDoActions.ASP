<HTML>
  <HEAD>
		<title>Ejecutar las acciones sobre un albaran de venta (un ticket)</title>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">
  </HEAD>
  <BODY>
 <%	Dim blnHayError
 	
 		blnHayError = false
 		
 		If Request.QueryString("Action")= vbNullString Then
 			Response.Write "Error, no se ha proporcionado valor al parámetro Action"	
 			blnHayError = True
 		End If	
 		
		' Creamos el objeto y recuperamos el estado de la variable de sesión.
   	Dim objAlbaranVenta
   	Dim strIncidencias
   	
   	strIncidencias = vbNullString
   	
   	Set objAlbaranVenta = Server.CreateObject("GescomObjects.AlbaranVenta")
		objAlbaranVenta.SetSuperState Session("AlbaranVentaState")
		

		Select Case Request.QueryString("Action")
		Case "AlbaranItemCodigoBarras"		' Dar de alta una nueva linea del ticket
	 		If Request.QueryString("txtBarCode")= vbNullString Then
	 			Response.Write "Error, no se ha proporcionado valor al parámetro txtBarCode"	
	 			blnHayError = True
	 		Else
	 			AlbaranItemCodigoBarras Request.QueryString("txtBarCode")
				Session("AlbaranVentaState") = objAlbaranVenta.GetSuperState
				Response.Redirect("TPVEdit.ASP")
	 		End If	

	 		
		Case "EliminarLinea"		' Eliminar una linea introducida previamente
	 		If Request.QueryString("Key")= vbNullString Then
	 			Response.Write "Error, no se ha proporcionado valor al parámetro Key"	
	 			blnHayError = True
	 		Else
				objAlbaranVenta.AlbaranVentaItems.Remove CInt(Request.QueryString("Key"))
				Session("AlbaranVentaState") = objAlbaranVenta.GetSuperState
				Response.Redirect("TPVEdit.ASP")
 			End If	
	
		Case "ModificarLinea"		' Modificar una linea introducida previamente
	 		If Request.QueryString("Key")= vbNullString Then
	 			Response.Write "Error, no se ha proporcionado valor al parámetro Key"	
	 			blnHayError = True
	 		ElseIf Request.QueryString("Descuento")= vbNullString Then
	 			Response.Write "Error, no se ha proporcionado valor al parámetro Descuento"	
	 			blnHayError = True
	 		ElseIf Request.QueryString("ImporteBruto")= vbNullString Then
	 			Response.Write "Error, no se ha proporcionado valor al parámetro ImporteBruto"	
	 			blnHayError = True
	 		Else
	 			Dim objAlbaranVentaItem
	 			Dim dblImporteBruto
	 			Dim dblDescuento
	 			
	 			Set objAlbaranVentaItem = objAlbaranVenta.AlbaranVentaItems(CInt(Request.QueryString("Key")))
	 			objAlbaranVentaItem.BeginEdit

				
	 			dblImporteBruto = CDbl(Request.QueryString("ImporteBruto"))
	 			dblDescuento = CDbl(Request.QueryString("Descuento"))
	 			'Si se ha modificado el precio de venta--> asumimos que es el definitivo y calculamos el importe equivalente
	 			If objAlbaranVentaItem.Bruto <> dblImporteBruto Then
	 				objAlbaranVentaItem.Bruto = Request.QueryString("ImporteBruto")
	 				objAlbaranVentaItem.CalcularDescuentoEquivalente

	 			ElseIf objAlbaranVentaItem.Descuento <> dblDescuento Then
	 				objAlbaranVentaItem.Descuento = Request.QueryString("Descuento")

	 			End If
	 			objAlbaranVentaItem.ApplyEdit
	 			Session("AlbaranVentaState") = objAlbaranVenta.GetSuperState
				Response.Redirect("TPVEdit.ASP")
 			End If	
	
		Case "ApplyEdit"		' Crear el registro del ticket de venta
	 		If Request.QueryString("Cliente")<> vbNullString Then
	 			objAlbaranVenta.Cliente = Request.QueryString("Cliente")
	 		End If
	 		If Request.QueryString("Representante")<> vbNullString Then
	 			objAlbaranVenta.Representante = Request.QueryString("Representante")
	 		End If
	 		If Request.QueryString("FormaPago")<> vbNullString Then
	 			objAlbaranVenta.FormaPago = Request.QueryString("FormaPago")
	 		End If
	 		If Request.QueryString("MedioPago")<> vbNullString Then
	 			objAlbaranVenta.MedioPago = Request.QueryString("MedioPago")
	 		End If
	 		If objAlbaranVenta.IsNew And Request.QueryString("Fecha")<> vbNullString Then
	 			objAlbaranVenta.Fecha = Request.QueryString("Fecha")
	 		End If
	   		objAlbaranVenta.AltaTicketTPV
			Session("AlbaranVentaState") = vbNullString
			Session("ImpresionTicket") = objAlbaranVenta.ImpresionTicket
	   		set objAlbaranVenta = Nothing
			Response.Redirect("TPVEditPrintTicket.ASP")

		Case "CancelEdit"		' Cancelar la edición del registro del ticket de venta y comenzar una nueva edición
	   		objAlbaranVenta.CancelEdit
			Session("AlbaranVentaState") = vbNullString
	   	set objAlbaranVenta = Nothing
			Response.Redirect("TPVEdit.ASP")

		Case "PrintLastTicket"		' Imprimir el último ticket
			Response.Redirect("TPVEditPrintTicket.ASP")

		Case "AlternarDevolucion"		' Cambia el modo de introducir lineas de normal a devolucion y viceversa
			objAlbaranVenta.EsDevolucion = not objAlbaranVenta.EsDevolucion
			Session("AlbaranVentaState") = objAlbaranVenta.GetSuperState
			Response.Redirect("TPVEdit.ASP")
		
	 	Case Else									' Dar error
 			Response.Write "Error, no existe la accion solicitada:" & Request.QueryString("Action")
 			blnHayError = True
					
			
		End Select

 		
 %>
  </BODY>
</HTML>
 <%
 		Sub AlbaranItemCodigoBarras (txtBarCode)
 			Dim lngCodigo
 			Dim intTalla
 			Dim strIncidencia
 			
			 ' Validar que tenga al menos información de talla + articulo
			lngCodigo = CLng(txtBarCode)
			If lngCodigo < 100 Then
	 				Response.Write "Falta información del código de artículo!, " & txtBarCode	
	 				blnHayError = True
			    Exit Sub
			End If
        
			' Validar que la información de talla sea correcta
			intTalla = CInt(Left(txtBarCode, 2))
			If intTalla Mod 2 <> 0 Then
	 				Response.Write "Talla errónea! (" & intTalla & "), en el código " & txtBarCode
	 				blnHayError = True
			    Exit Sub
			End If
			If intTalla > 56 Or intTalla < 36 Then
			    Response.Write "Talla errónea! (" & intTalla & "), en el código " & txtBarCode
	 				blnHayError = True
			    Exit Sub
			End If
        
			strIncidencia = objAlbaranVenta.AlbaranItemTPV(CStr(txtBarCode))
			Response.Write txtBarCode
			If strIncidencia <> vbNullString Then
			    Response.Write  strIncidencia & " Código:" & txtBarCode
	 				blnHayError = True
	 				Exit Sub
			End If
        
        
		End Sub
%>