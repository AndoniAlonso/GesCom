<HTML>
  <HEAD>
		<title>Salvar las configuración de las cookies</title>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">
  </HEAD>
  <BODY>
 <%	Dim blnHayError
 	
 		blnHayError = false
 		
 		If Request.QueryString("TerminalID")= vbNullString Then
 			Response.Write "Error, no se ha proporcionado valor al parámetro TerminalID"	
 			blnHayError = True
 		Else
	 		Response.Cookies("TerminalID") = Request.QueryString("TerminalID")
	 		Response.Cookies("TerminalID").Expires = "31/12/2037"
 		End If	
 	
 		If Request.QueryString("CentroGestionID")= vbNullString Then
 			Response.Write "Error, no se ha proporcionado valor al parámetro CentroGestionID"
 			blnHayError = True
 		Else
	 		Response.Cookies("CentroGestionID") = Request.QueryString("CentroGestionID")
	 		Response.Cookies("CentroGestionID").Expires = "31/12/2037"
 		End If	

 		If Request.QueryString("AlmacenID")= vbNullString Then
 			Response.Write "Error, no se ha proporcionado valor al parámetro AlmacenID"
 			blnHayError = True
 		Else
	 		Response.Cookies("AlmacenID") = Request.QueryString("AlmacenID")
	 		Response.Cookies("AlmacenID").Expires = "31/12/2037"
 		End If	


		If not blnHayError Then
			Response.Redirect("TPVEdit.ASP")
		End If
 		
 %>
  </BODY>