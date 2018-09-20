<HTML>
  <HEAD>
		<title>Gescom TPV</title>
		<LINK href="./styles/GescomTPV.css" type="text/css" rel="stylesheet">
  </HEAD>
  <BODY>
<table id="Container" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
  <tr>
	 <td vAlign="top" width="20%">&nbsp;
   </td>
	<td width="3%">&nbsp;</td>
	<td vAlign="top">
		<table cellSpacing="0" cellPadding="0" width="540" border="0">
			<tr>
				<td class="cadre_titre">Seleccionar el terminal</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr style="color:White;background-color:#6B696B;font-weight:bold;">
				<td align='center'>Terminal</td><td>Nombre</td><td>Centro de gestión</td><td>Almacén</td>
			</tr>
			<% Dim rcsTerminales
				
				Set rcsTerminales = Server.CreateObject("ADODB.Recordset")
				
				rcsTerminales.Open "Select TOP 100 TerminalID, Nombre, CentroGestionID, AlmacenID FROM vTerminales", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=GestionComercial;Data Source=localhost"
				
				Do While Not rcsTerminales.EOF
			%>   <tr nowrap="nowrap" onmouseover="this.style.backgroundColor='powderblue'; this.style.cursor='hand';" onmouseout="this.style.backgroundColor='#F7F7DE';" style="background-color:#F7F7DE;">
			<%   Response.Write "<TD align='center'>"
				  Response.Write rcsTerminales("TerminalID")
				  Response.Write "</TD>"
				  Response.Write "<TD>"
				  Response.Write "<A href='TPVIdentifySaveCookies.asp?TerminalID=" & rcsTerminales("TerminalID") & "&CentroGestionID=" & rcsTerminales("CentroGestionID") & "&AlmacenID=" & rcsTerminales("AlmacenID") & "'>"
				  Response.Write rcsTerminales("Nombre")
				  Response.Write "</A>"
				  Response.Write "</TD>"
				  Response.Write "<TD align='center'>"
				  Response.Write rcsTerminales("CentroGestionID")
				  Response.Write "</TD>"
				  Response.Write "<TD align='center'>"
				  Response.Write rcsTerminales("AlmacenID")
				  Response.Write "</TD>"
				  Response.Write "</TR>"
				
				  rcsTerminales.MoveNext
				Loop
				
				rcsTerminales.Close
				Set rcsTerminales = Nothing
			%>
		</table>
	</td>
</tr>
</table>
		
</BODY>
</HTML>
