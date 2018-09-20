		<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr> 
				<td class="cadre_titre"><%
				If session("NombreCentroGestion")= vbNullString Then
					Dim objCentroGestion 
					Set objCentroGestion = Server.CreateObject("GescomObjects.CentroGestion")
					
					objCentroGestion.Load Request.Cookies("CentroGestionID")
					Session("NombreCentroGestion") = objCentroGestion.Nombre 
					Set objCentroGestion = Nothing
				End If
				Response.Write Session("NombreCentroGestion")
				%></td>
		  </tr>
		  <tr> 
		    <td class="cadre_titre">Opciones de TPV</td>
		  </tr>
		  <tr> 
		    <td>&nbsp;</td>
		
		  </tr>
		  <tr> 
		      <td class="cadre_menu"><a href="TPVEdit.ASP">Punto de venta</a></td>
		  </tr>
	     <tr> 
		    <td>&nbsp;</td>
		  </tr>
		  <tr> 
		    <td class="cadre_menu"><a href="javascript:RedirectTo('CriteriaCierreCaja.ASP');">Cierre de caja</a></td>
		  </tr>
	     <tr> 
		    <td>&nbsp;</td>
		  </tr>
		  <tr> 
		    <td class="cadre_menu"><a href="javascript:RedirectTo('CriteriaListadoVentas.ASP');">Listado de ventas</a></td>
		  </tr>
		</table>
